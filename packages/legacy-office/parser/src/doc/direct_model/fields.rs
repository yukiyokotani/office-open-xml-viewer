//! Direct projection of Word fields onto the DOCX model's field semantics.
//!
//! MS-DOC 2.8.25 (Plcfld), 2.9.88 (Fld), 2.9.89 (fldch), 2.9.90 (flt) and
//! 2.9.110 (grffldEnd); ECMA-376 Part 1 17.16.
//!
//! The shared tokenizer already hides every field instruction and keeps every
//! stored result (the content between the separator and the end character).
//! This module validates a story's field characters against that story's own
//! Plcfld and then decides, for every visible field, what Word displays:
//!
//! * PAGE, NUMPAGES, DATE and TIME are the fields the DOCX parser turns into
//!   renderer-evaluated `DocRun::Field` values (`classify_field` in the DOCX
//!   parser). Word re-evaluates them for display: the Word PDF of a DOC whose
//!   header DATE field stores an older date shows the export date, and page
//!   numbers follow pagination. They are projected onto the same `FieldRun`.
//!   A switch the shared renderer does not interpret leaves the value in its
//!   default format, or the stored result for a date, as it does for a DOCX.
//!   A locked field (grffldEnd fLocked) is not updated: it displays its
//!   stored result. The dirty and edited flags do not stop Word from
//!   updating the field when it prints (Word controls).
//! * Every other field displays its stored result, which stays ordinary runs,
//!   matching the DOCX parser's complex-field behavior. Fields never execute.
//! * EQ and a SYMBOL without a stored result are drawn by Word itself: an EQ
//!   field becomes a phonetic guide or an inline equation (`equation.rs`),
//!   a SYMBOL field one character in its font (ECMA-376 17.16.5.63).
//! * MACROBUTTON, GOTOBUTTON, ADVANCE and FORMDROPDOWN are rejected until
//!   they are projected.
//! * grffldEnd fDiffer inverts the field-code view of one field. Word prints
//!   the result of such a field (Word control), so the flag changes nothing.
//!
//! A field result flagged fPrivateResult "is not intended to be visible to
//! the user" (MS-DOC 2.9.110): its content is not displayed. INCLUDEPICTURE
//! is the exception: the Word PDF of a DOC with such a picture result shows
//! the picture.

use crate::doc::{header_fields, unsupported, Paragraph, Token};

/// Maximum retained instruction/result UTF-8 bytes per field. The value only
/// bounds scratch work: a longer instruction of an evaluated field is
/// rejected, while a longer cached-field instruction only needs its keyword.
const MAX_FIELD_TEXT: usize = 1024;
/// Resource policy for open field frames, not an MS-DOC limit.
const MAX_FIELD_DEPTH: usize = 256;

/// One renderer-evaluated field (ECMA-376 17.16.5.13/42/44/65).
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::doc) struct Evaluated {
    /// The DOCX `FieldRun.field_type`: `page`, `numPages`, `date` or `time`.
    pub(in crate::doc) field_type: &'static str,
    /// Trimmed instruction, as the DOCX parser stores `FieldRun.instruction`.
    pub(in crate::doc) instruction: String,
    /// Stored result text, the DOCX `FieldRun.fallback_text`.
    pub(in crate::doc) cached_result: String,
    /// Character whose properties format the evaluated result.
    pub(in crate::doc) format_cp: usize,
    /// First stored result character when its formatting must agree with
    /// `format_cp` (no general-formatting switch selects the source).
    pub(in crate::doc) agreeing_result_cp: Option<usize>,
    /// FORMCHECKBOX only: the instruction's binary-data character, whose
    /// NilPICFAndBinData holds the FFData (MS-DOC 2.9.78, 2.9.158).
    pub(in crate::doc) form_data_cp: Option<usize>,
    /// What Word draws for a field that is not a `FieldRun`.
    pub(in crate::doc) drawn: Option<Box<Drawn>>,
}

pub(in crate::doc) use super::equation::RubyForm;

/// The display of a field Word draws itself.
#[derive(Debug, Clone)]
pub(in crate::doc) enum Drawn {
    /// EQ phonetic guide.
    Ruby(RubyForm),
    /// Any other EQ field, as an inline equation.
    Equation(Vec<ooxml_common::math::MathNode>),
    /// SYMBOL without a stored result.
    Symbol(SymbolForm),
}

impl PartialEq for Drawn {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Ruby(a), Self::Ruby(b)) => a == b,
            (Self::Symbol(a), Self::Symbol(b)) => a == b,
            // The shared equation model has no equality of its own.
            (Self::Equation(a), Self::Equation(b)) => {
                serde_json::to_value(a).ok() == serde_json::to_value(b).ok()
            }
            _ => false,
        }
    }
}

impl Eq for Drawn {}

/// ECMA-376 17.16.5.63 SYMBOL: one character, by code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::doc) struct SymbolForm {
    pub(in crate::doc) code: u32,
    /// `\f`: the font the character is drawn in.
    pub(in crate::doc) font: Option<String>,
    /// `\s`: the font size in half-points.
    pub(in crate::doc) size_half_points: Option<u16>,
    /// `\u`: the code is a Unicode scalar value, not an ANSI code.
    pub(in crate::doc) unicode: bool,
}

/// Link semantics of stored-result content, as the DOCX parser derives them
/// for runs inside `<w:hyperlink>` (ECMA-376 17.16.22) and inside REF/PAGEREF
/// results with the `\h` switch (17.16.5.45/51). A DOC HYPERLINK field
/// (17.16.5.25) is the field form of `<w:hyperlink>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::doc) struct Link {
    /// External target (`TextRun.hyperlink`).
    pub(in crate::doc) href: Option<String>,
    /// Bookmark target (`TextRun.hyperlink_anchor`).
    pub(in crate::doc) anchor: Option<String>,
    /// Inside a TOC field result: the DOCX parser keeps the paragraph-level
    /// color and underline for link runs there, as Word displays TOC entries
    /// (observed: a DOC TOC whose HYPERLINK results carry the blue underlined
    /// Hyperlink character style prints in the TOC paragraph style).
    pub(in crate::doc) in_toc: bool,
}

/// A displayed token together with its link semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::doc) struct Linked {
    pub(in crate::doc) token: Token,
    pub(in crate::doc) link: Link,
}

/// Validated field structure of one aggregate story (main, header, footnote
/// or endnote document), with CPs relative to that document.
pub(in crate::doc) struct StoryFields {
    evaluated: Vec<(usize, usize, Evaluated)>,
    /// Piecewise-constant link context: each entry applies from its CP up to
    /// the next entry.
    links: Vec<(usize, Option<Link>)>,
    /// Sorted, disjoint CP ranges (first, last) of private results, whose
    /// content is not displayed.
    private: Vec<(usize, usize)>,
}

/// What `classify` decides for one field.
enum Display {
    /// The stored result, as ordinary runs.
    Result,
    /// One token in place of the field.
    Evaluated(Evaluated),
    /// Nothing: the stored result is private.
    Private,
}

#[derive(Debug, Default, Clone, PartialEq, Eq)]
enum Role {
    #[default]
    None,
    Toc,
    Hyperlink {
        href: Option<String>,
        anchor: Option<String>,
    },
    Anchor(String),
}

#[derive(Default)]
struct Open {
    begin: usize,
    parent: bool,
    listed_type: Option<u8>,
    hidden: bool,
    role: Role,
    separator: Option<usize>,
    instruction: String,
    instruction_overflow: bool,
    instruction_controls: bool,
    form_data_cp: Option<usize>,
    /// Instruction characters with their CPs (bounded like `instruction`).
    instruction_chars: Vec<(char, usize)>,
    keyword_cp: Option<usize>,
    child_in_instruction: bool,
    child_in_result: bool,
    result: String,
    result_overflow: bool,
    result_cp: Option<usize>,
    result_controls: bool,
}

impl StoryFields {
    /// Validate `text` (the complete story) against its Plcfld and select the
    /// fields Word evaluates. `partitions` are sorted CPs that separate
    /// independently projected stories inside the aggregate document; no
    /// field may span one of them.
    pub(in crate::doc) fn analyze(
        text: &str,
        table: &header_fields::Table,
        partitions: &[usize],
    ) -> Result<Self, String> {
        let mut stack: Vec<Open> = Vec::new();
        let mut evaluated: Vec<(usize, usize, Evaluated)> = Vec::new();
        let mut private: Vec<(usize, usize)> = Vec::new();
        let mut links = Vec::new();
        let mut listed = 0usize;
        let mut cp = 0usize;
        for character in text.chars() {
            match character {
                '\u{13}' => {
                    let listed_type = match table.get(cp) {
                        Some((0x13, kind)) => {
                            listed += 1;
                            Some(kind)
                        }
                        None => None,
                        Some(_) => return Err(mismatch()),
                    };
                    let hidden = stack
                        .last_mut()
                        .map(|parent| {
                            if parent.separator.is_none() {
                                parent.child_in_instruction = true;
                            } else {
                                parent.child_in_result = true;
                            }
                            parent.hidden || parent.separator.is_none()
                        })
                        .unwrap_or(false);
                    if stack.len() >= MAX_FIELD_DEPTH {
                        return Err(crate::rejection::resource(
                            "Word field nesting budget exceeded",
                        ));
                    }
                    stack.push(Open {
                        begin: cp,
                        parent: !stack.is_empty(),
                        listed_type,
                        hidden,
                        ..Open::default()
                    });
                }
                '\u{14}' => {
                    let field = stack
                        .last_mut()
                        .ok_or_else(|| unsupported("unbalanced Word field separator"))?;
                    match (field.listed_type.is_some(), table.get(cp)) {
                        (true, Some((0x14, _))) => listed += 1,
                        (false, None) => {}
                        _ => return Err(mismatch()),
                    }
                    if field.separator.replace(cp).is_some() {
                        return Err(unsupported("Word field has two separators"));
                    }
                    if !field.hidden && field.listed_type.is_some() {
                        field.role = role(&field.instruction);
                        if field.role != Role::None {
                            links.push((cp + 1, effective_link(&stack)));
                        }
                    }
                }
                '\u{15}' => {
                    let field = stack
                        .pop()
                        .ok_or_else(|| unsupported("unbalanced Word field end"))?;
                    let flags = match (field.listed_type.is_some(), table.get(cp)) {
                        (true, Some((0x15, flags))) => {
                            listed += 1;
                            Some(flags)
                        }
                        (false, None) => None,
                        _ => return Err(mismatch()),
                    };
                    let first = partitions.partition_point(|boundary| *boundary <= field.begin);
                    let last = partitions.partition_point(|boundary| *boundary <= cp);
                    if first != last {
                        return Err(unsupported("Word field crosses a story boundary"));
                    }
                    match classify(&field, flags)? {
                        Display::Result => {}
                        Display::Evaluated(entry) => {
                            // The token replaces everything the field holds,
                            // fields nested in its result included. Those
                            // closed after this field began: they are the
                            // tail of the list.
                            while evaluated
                                .last()
                                .is_some_and(|(begin, ..)| *begin > field.begin)
                            {
                                evaluated.pop();
                            }
                            evaluated.push((field.begin, cp, entry));
                        }
                        Display::Private => {
                            if let Some(separator) = field.separator {
                                while evaluated
                                    .last()
                                    .is_some_and(|(begin, ..)| *begin > field.begin)
                                {
                                    evaluated.pop();
                                }
                                while private
                                    .last()
                                    .is_some_and(|(first, _)| *first > field.begin)
                                {
                                    private.pop();
                                }
                                private.push((separator + 1, cp));
                            }
                        }
                    }
                    if field.role != Role::None {
                        links.push((cp, effective_link(&stack)));
                    }
                }
                _ => {
                    if let Some(field) = stack.last_mut() {
                        field.push(character, cp);
                    }
                }
            }
            cp += character.len_utf16();
        }
        if !stack.is_empty() {
            return Err(unsupported("unbalanced Word field begin"));
        }
        if listed != table.len() {
            return Err(unsupported(
                "Word field table lists characters absent from its story",
            ));
        }
        // A field closing after the fields nested in it dropped them, so the
        // remaining ranges are disjoint.
        evaluated.sort_unstable_by_key(|(begin, ..)| *begin);
        private.sort_unstable();
        Ok(Self {
            evaluated,
            links,
            private,
        })
    }

    fn is_private(&self, cp: usize) -> bool {
        let index = self.private.partition_point(|(first, _)| *first <= cp);
        index
            .checked_sub(1)
            .is_some_and(|index| cp <= self.private[index].1)
    }

    fn link_at(&self, cp: usize) -> Option<&Link> {
        let index = self.links.partition_point(|(start, _)| *start <= cp);
        index
            .checked_sub(1)
            .and_then(|index| self.links[index].1.as_ref())
    }

    /// Replace each evaluated field's stored-result tokens in one already
    /// tokenized slice (`base_cp` is the slice's first CP) with one
    /// `Token::EvaluatedField`, placed at the field's begin character.
    pub(in crate::doc) fn apply(
        &self,
        base_cp: usize,
        paragraphs: &mut [Paragraph],
    ) -> Result<(), String> {
        let Some(limit) = paragraphs.last().map(|paragraph| paragraph.end_cp) else {
            return Ok(());
        };
        let first = self
            .evaluated
            .partition_point(|(begin, ..)| *begin < base_cp);
        let mut pending = self.evaluated[first..]
            .iter()
            .take_while(|(begin, ..)| *begin <= limit)
            .peekable();
        let mut covered_until = None;
        for paragraph in paragraphs {
            for (token, cp) in std::mem::take(&mut paragraph.tokens) {
                while let Some((begin, end, field)) = pending.next_if(|(begin, ..)| *begin < cp) {
                    paragraph
                        .tokens
                        .push((Token::EvaluatedField(Box::new(field.clone())), *begin));
                    covered_until = Some(*end);
                }
                if covered_until.is_some_and(|end| cp <= end) || self.is_private(cp) {
                    // The stored result Word replaces when it evaluates the
                    // field, or a result it does not display.
                    continue;
                }
                paragraph.tokens.push((token, cp));
            }
            while let Some((begin, end, field)) =
                pending.next_if(|(begin, ..)| *begin <= paragraph.end_cp)
            {
                paragraph
                    .tokens
                    .push((Token::EvaluatedField(Box::new(field.clone())), *begin));
                covered_until = Some(*end);
            }
            if self.links.is_empty() {
                continue;
            }
            for (token, cp) in &mut paragraph.tokens {
                // A FieldRun carries no link in the DOCX model.
                if matches!(token, Token::EvaluatedField(_)) {
                    continue;
                }
                if let Some(link) = self.link_at(*cp) {
                    let inner = std::mem::replace(token, Token::Tab);
                    *token = Token::Linked(Box::new(Linked {
                        token: inner,
                        link: link.clone(),
                    }));
                }
            }
        }
        Ok(())
    }
}

impl Open {
    fn push(&mut self, character: char, cp: usize) {
        if self.separator.is_none() {
            if character < ' ' && character != '\t' {
                if character == '\u{1}' && self.form_data_cp.is_none() {
                    self.form_data_cp = Some(cp);
                } else {
                    self.instruction_controls = true;
                }
                return;
            }
            if self.instruction_chars.len() < MAX_FIELD_TEXT {
                self.instruction_chars.push((character, cp));
            }
            if self.keyword_cp.is_none() && !character.is_whitespace() {
                self.keyword_cp = Some(cp);
            }
            if self.instruction.len() < MAX_FIELD_TEXT {
                self.instruction.push(character);
            } else {
                self.instruction_overflow = true;
            }
        } else {
            self.result_cp.get_or_insert(cp);
            if character < ' ' {
                self.result_controls = true;
            } else if self.result.len() < MAX_FIELD_TEXT {
                self.result.push(character);
            } else {
                self.result_overflow = true;
            }
        }
    }

    fn has_result_content(&self) -> bool {
        self.result_cp.is_some() || self.child_in_result
    }
}

/// DOCX precedence: a hyperlink's own target and anchor, else the innermost
/// `\h` field anchor; any enclosing TOC result marks the run as a TOC link.
fn effective_link(stack: &[Open]) -> Option<Link> {
    let mut in_toc = false;
    let mut hyperlink = None;
    let mut field_anchor = None;
    for field in stack {
        match &field.role {
            Role::None => {}
            Role::Toc => in_toc = true,
            Role::Hyperlink { href, anchor } => {
                // A run has one target: the innermost hyperlink, the one
                // Word follows.
                if hyperlink.replace((href, anchor)).is_some() {
                    crate::rejection::approximated(
                        "nested Word hyperlink fields link to the innermost target",
                        (),
                    );
                }
            }
            Role::Anchor(anchor) => field_anchor = Some(anchor),
        }
    }
    match hyperlink {
        Some((href, anchor)) => Some(Link {
            href: href.clone(),
            anchor: anchor.clone().or_else(|| field_anchor.cloned()),
            in_toc,
        }),
        None => field_anchor.map(|anchor| Link {
            href: None,
            anchor: Some(anchor.clone()),
            in_toc,
        }),
    }
}

/// Result-display role of a visible cached field, from its instruction.
fn role(instruction: &str) -> Role {
    let words: Vec<&str> = instruction.split_whitespace().collect();
    let keyword = words.first().copied().unwrap_or_default();
    if keyword.eq_ignore_ascii_case("TOC") {
        return Role::Toc;
    }
    if keyword.eq_ignore_ascii_case("HYPERLINK") {
        return hyperlink(instruction);
    }
    // The DOCX parser's classify_complex_field: REF/PAGEREF with a `\h`
    // switch after the bookmark argument link to that bookmark.
    if (keyword.eq_ignore_ascii_case("REF") || keyword.eq_ignore_ascii_case("PAGEREF"))
        && words
            .iter()
            .skip(2)
            .any(|word| word.eq_ignore_ascii_case("\\h"))
    {
        if let Some(anchor) = words
            .get(1)
            .map(|target| target.trim_matches(['\'', '"']))
            .filter(|target| !target.is_empty())
        {
            return Role::Anchor(anchor.to_string());
        }
    }
    Role::None
}

/// ECMA-376 17.16.5.25 HYPERLINK with 17.16.1 argument quoting: an optional
/// target argument, `\l` location, and the display-neutral `\m`, `\n`,
/// `\o` and `\t` switches. An instruction that does not read this way links
/// nowhere; its result is displayed like any other.
fn hyperlink(instruction: &str) -> Role {
    fn read(instruction: &str) -> Option<Role> {
        let items = arguments(instruction)?;
        let mut items = items.into_iter().skip(1).peekable();
        let href = items.next_if(|(switch, _)| !switch).map(|(_, value)| value);
        let mut anchor = None;
        while let Some((switch, name)) = items.next() {
            if !switch {
                return None;
            }
            match name.as_str() {
                "\\l" | "\\o" | "\\t" => {
                    let (_, value) = items.next_if(|(switch, _)| !switch)?;
                    if name == "\\l" && anchor.replace(value).is_some() {
                        return None;
                    }
                }
                // Word also writes `\h` on HYPERLINK fields; like `\m` and `\n`
                // it takes no argument and does not change the stored result.
                "\\m" | "\\n" | "\\h" => {}
                _ => return None,
            }
        }
        Some(Role::Hyperlink { href, anchor })
    }
    read(instruction).unwrap_or_else(|| {
        crate::rejection::skipped("Word hyperlink instruction without a readable target");
        Role::None
    })
}

/// Split an instruction into (is_switch, text) items. Quoted arguments may
/// contain white space; `\"` and `\\` escape a quote and a backslash
/// (ECMA-376 17.16.1). `None` for an unterminated quote.
fn arguments(instruction: &str) -> Option<Vec<(bool, String)>> {
    let mut items = Vec::new();
    let mut characters = instruction.chars().peekable();
    while let Some(&character) = characters.peek() {
        if character.is_whitespace() {
            characters.next();
            continue;
        }
        if character == '"' {
            characters.next();
            let mut value = String::new();
            loop {
                match characters.next()? {
                    '"' => break,
                    '\\' if matches!(characters.peek(), Some('"' | '\\')) => {
                        value.push(characters.next()?)
                    }
                    other => value.push(other),
                }
            }
            items.push((false, value));
            continue;
        }
        let mut value = String::new();
        while let Some(&next) = characters.peek() {
            if next.is_whitespace() {
                break;
            }
            value.push(next);
            characters.next();
        }
        items.push((value.starts_with('\\'), value));
    }
    Some(items)
}

fn mismatch() -> String {
    unsupported("Word field characters disagree with the field table")
}

/// Field types whose field characters MS-DOC 2.8.25 forbids in a Plcfld. Word
/// stores them in the text only; none of them displays a result.
const UNLISTED_TYPES: [&str; 5] = ["XE", "TC", "RD", "TA", "PRIVATE"];

/// MS-DOC 2.9.110 grffldEnd.
const F_LOCKED: u8 = 0x10;
const F_PRIVATE_RESULT: u8 = 0x20;

fn classify(field: &Open, flags: Option<u8>) -> Result<Display, String> {
    let keyword = field
        .instruction
        .split_whitespace()
        .next()
        .unwrap_or_default()
        .to_ascii_uppercase();
    let Some(flags) = flags else {
        if !UNLISTED_TYPES.contains(&keyword.as_str()) || field.separator.is_some() {
            return Err(mismatch());
        }
        return Ok(Display::Result);
    };
    // MS-DOC 2.9.110: fHasSep and fNested MUST describe the actual structure.
    if (flags & 0x80 != 0) != field.separator.is_some() || (flags & 0x40 != 0) != field.parent {
        return Err(unsupported("Word field flags disagree with its structure"));
    }
    if field.hidden {
        return Ok(Display::Result);
    }
    // A field Word evaluates or draws is read from its whole instruction and
    // replaced as a whole: scratch that did not hold all of it is a limit of
    // this reader, not content to drop.
    let evaluated_keyword = matches!(
        keyword.as_str(),
        "PAGE" | "NUMPAGES" | "DATE" | "TIME" | "EQ" | "SYMBOL"
    );
    if evaluated_keyword && (field.instruction_overflow || field.result_overflow) {
        return Err(crate::rejection::resource(
            "Word evaluated field text budget exceeded",
        ));
    }
    // fDiffer (0x01) only inverts the field-code view on screen; fZombieEmbed
    // (0x02), fResultsDirty (0x04) and fResultsEdited (0x08) do not change
    // what Word prints: it updates the field (Word controls: an edited PAGE
    // result prints the page number).
    let stored = || {
        if flags & F_PRIVATE_RESULT != 0
            && field.has_result_content()
            && keyword != "INCLUDEPICTURE"
        {
            Display::Private
        } else {
            Display::Result
        }
    };
    let (field_type, listed_type) = match keyword.as_str() {
        "PAGE" => ("page", 0x21),
        "NUMPAGES" => ("numPages", 0x1a),
        "DATE" => ("date", 0x1f),
        "TIME" => ("time", 0x20),
        "FORMCHECKBOX" => return checkbox(field, flags).map(Display::Evaluated),
        "EQ" => return Ok(equation(field)),
        "MACROBUTTON" | "GOTOBUTTON" | "ADVANCE" | "FORMDROPDOWN" => {
            return Err(unsupported(format!(
                "Word {keyword} field display is not supported"
            )));
        }
        "SYMBOL" if field.separator.is_none() => return Ok(symbol(field)),
        _ => return Ok(stored()),
    };
    if field.listed_type != Some(listed_type) {
        return Err(unsupported(
            "Word evaluated field type disagrees with its instruction",
        ));
    }
    // A locked field is not updated: Word displays the stored result.
    if flags & F_LOCKED != 0 {
        return Ok(stored());
    }
    let (instruction, source) = match field_type {
        "page" | "numPages" => number_switches(&field.instruction),
        _ => (
            field.instruction.trim().to_string(),
            date_switches(&field.instruction),
        ),
    };
    let instruction_cp = field.begin + 1;
    let (format_cp, agreeing_result_cp) = match source {
        // ECMA-376 17.16.4.3.3: CHARFORMAT formats the result like the first
        // character of the field-type name.
        FormatSource::CharFormat => (field.keyword_cp.unwrap_or(instruction_cp), None),
        // MERGEFORMAT keeps the stored result's formatting. The DOCX parser
        // takes it from the first stored-result run.
        FormatSource::MergeFormat => (field.result_cp.unwrap_or(instruction_cp), None),
        // Without a switch the DOCX parser formats the value like the first
        // instruction run. Word's choice is not specified; the projector
        // records a stored result formatted differently.
        FormatSource::Default => (instruction_cp, field.result_cp),
    };
    Ok(Display::Evaluated(Evaluated {
        field_type,
        instruction,
        cached_result: field.result.clone(),
        format_cp,
        agreeing_result_cp,
        form_data_cp: None,
        drawn: None,
    }))
}

/// An EQ field stores no result: Word draws the instruction (`equation.rs`).
fn equation(field: &Open) -> Display {
    let (drawn, format_cp) = match super::equation::parse(&field.instruction_chars) {
        super::equation::Equation::Ruby(form) => {
            let cp = form.base_cp;
            (Drawn::Ruby(form), cp)
        }
        super::equation::Equation::Math(nodes, cp) => {
            (Drawn::Equation(nodes), cp.unwrap_or(field.begin + 1))
        }
    };
    Display::Evaluated(Evaluated {
        field_type: if matches!(drawn, Drawn::Ruby(_)) {
            "ruby"
        } else {
            "equation"
        },
        instruction: field.instruction.trim().to_string(),
        cached_result: String::new(),
        format_cp,
        agreeing_result_cp: None,
        form_data_cp: None,
        drawn: Some(Box::new(drawn)),
    })
}

/// ECMA-376 17.16.5.63 SYMBOL: "the character whose code point value is
/// specified", decimal or `0x` hexadecimal, with `\f` font, `\s` size and
/// the `\a` (ANSI, the default), `\u` (Unicode) and `\j` (Shift-JIS) code
/// interpretations. `\h` does not change the character.
fn symbol(field: &Open) -> Display {
    let items = arguments(&field.instruction).unwrap_or_default();
    let mut items = items.into_iter().skip(1).peekable();
    let code = items.next_if(|(switch, _)| !switch).and_then(|(_, value)| {
        match value
            .strip_prefix("0x")
            .or_else(|| value.strip_prefix("0X"))
        {
            Some(hex) => u32::from_str_radix(hex, 16).ok(),
            None => value.parse().ok(),
        }
    });
    let (mut font, mut size_half_points, mut unicode) = (None, None, false);
    while let Some((switch, name)) = items.next() {
        if !switch {
            continue;
        }
        match name.to_ascii_lowercase().as_str() {
            "\\f" => font = items.next_if(|(switch, _)| !switch).map(|(_, value)| value),
            "\\s" => {
                // MS-DOC 2.6.1 sprmCHps: a size of 1 to 1638 points.
                size_half_points = items
                    .next_if(|(switch, _)| !switch)
                    .and_then(|(_, value)| value.parse::<f64>().ok())
                    .filter(|size| (1.0..=1638.0).contains(size))
                    .map(|size| (size * 2.0).round() as u16);
            }
            "\\u" => unicode = true,
            "\\j" => {
                // The reader holds no Shift-JIS table.
                crate::rejection::approximated(
                    "Word SYMBOL Shift-JIS code read as a Unicode value",
                    (),
                );
                unicode = true;
            }
            _ => {}
        }
    }
    let Some(code) = code else {
        // No character is named: nothing is drawn.
        crate::rejection::ignored_value("Word SYMBOL field without a character code");
        return Display::Private;
    };
    Display::Evaluated(Evaluated {
        field_type: "symbol",
        instruction: field.instruction.trim().to_string(),
        cached_result: String::new(),
        // 17.16.5.63: the field's own character formatting applies.
        format_cp: field.keyword_cp.unwrap_or(field.begin + 1),
        agreeing_result_cp: None,
        form_data_cp: None,
        drawn: Some(Box::new(Drawn::Symbol(SymbolForm {
            code,
            font,
            size_half_points,
            unicode,
        }))),
    })
}

/// ECMA-376 17.16.5.20 FORMCHECKBOX, which the DOCX parser turns into a
/// `checkbox` FieldRun drawn as a ballot box in the begin run's formatting.
/// MS-DOC stores the state in the FFData of the binary-data character inside
/// the instruction; the story projector reads it. The result is empty and
/// private (grffldEnd fPrivateResult) in Word's output.
fn checkbox(field: &Open, flags: u8) -> Result<Evaluated, String> {
    if field.listed_type != Some(0x47) {
        return Err(unsupported(
            "Word evaluated field type disagrees with its instruction",
        ));
    }
    let mut words = field.instruction.split_whitespace();
    words.next();
    // MS-OE376 2.1.505: Word shows nothing for a nested check box, while
    // the DOCX parser omits it; neither is established for DOC.
    if flags & 0x1e != 0
        || words.next().is_some()
        || field.parent
        || field.child_in_instruction
        || field.child_in_result
        || field.instruction_controls
        || field.instruction_overflow
        || field.has_result_content()
    {
        return Err(unsupported(
            "Word FORMCHECKBOX field with unsupported state or content",
        ));
    }
    let form_data_cp = field
        .form_data_cp
        .ok_or_else(|| unsupported("Word FORMCHECKBOX field lacks its form data"))?;
    Ok(Evaluated {
        field_type: "checkbox",
        instruction: field.instruction.trim().to_string(),
        cached_result: String::new(),
        // The DOCX parser formats legacy check boxes like the begin run.
        format_cp: field.begin,
        agreeing_result_cp: None,
        form_data_cp: Some(form_data_cp),
        drawn: None,
    })
}

/// Check-box state from an FFData (MS-DOC 2.9.78/2.9.79): whether it is
/// checked and its explicit size in points (`None` when sized automatically).
pub(in crate::doc) fn checkbox_state(data: &[u8]) -> Result<(bool, Option<f64>), String> {
    let invalid = || unsupported("invalid Word check-box form data");
    let word = |offset: usize| -> Result<u16, String> {
        data.get(offset..offset + 2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .ok_or_else(invalid)
    };
    if data.get(..4) != Some(&[0xff; 4][..]) {
        return Err(invalid());
    }
    let bits = word(4)?;
    let (kind, result, automatic_size) = (bits & 3, (bits >> 2) & 0x1f, bits & 0x400 != 0);
    if kind != 1 || word(6)? != 0 {
        return Err(invalid());
    }
    let size = word(8)?;
    // xstzName: Xst (cch, cch UTF-16 units) plus a zero terminator, then wDef.
    let name = usize::from(word(10)?);
    if name > 20 || word(12 + name * 2)? != 0 {
        return Err(invalid());
    }
    let default = word(14 + name * 2)?;
    let checked = match (result, default) {
        (0, _) => false,
        (1, _) => true,
        // "Undefined checkboxes are treated as unchecked." Keep a checked
        // default closed until an Office control shows which state Word draws.
        (25, 0) => false,
        _ => {
            return Err(unsupported(
                "Word check box with an undefined state and checked default",
            ))
        }
    };
    let size = if automatic_size {
        None
    } else if (2..=3168).contains(&size) {
        Some(f64::from(size) / 2.0)
    } else {
        return Err(invalid());
    };
    Ok((checked, size))
}

#[derive(Debug, PartialEq, Eq)]
enum FormatSource {
    Default,
    CharFormat,
    MergeFormat,
}

/// ECMA-376 17.16.4.3.1 general numeric formats that the shared field renderer
/// maps one-to-one. The renderer matches them case-sensitively; Word does
/// not (Word control: `\* ROMAN` prints `I`), and takes the capitalization of
/// the result from the first letter of the argument.
const NUMBER_FORMATS: [(&str, &str, &str); 4] = [
    ("arabic", "Arabic", "Arabic"),
    ("arabicdash", "ArabicDash", "ArabicDash"),
    ("hex", "Hex", "Hex"),
    ("roman", "Roman", "roman"),
];

fn canonical_number_format(argument: &str) -> Option<&'static str> {
    let lower = argument.to_ascii_lowercase();
    let upper = argument.starts_with(|first: char| first.is_ascii_uppercase());
    if lower == "alphabetic" {
        return Some(if upper { "ALPHABETIC" } else { "alphabetic" });
    }
    NUMBER_FORMATS
        .iter()
        .find(|(name, ..)| *name == lower)
        .map(|(_, capital, small)| if upper { *capital } else { *small })
}

fn general_format(argument: &str, source: &mut FormatSource) -> bool {
    let selected = if argument.eq_ignore_ascii_case("MERGEFORMAT") {
        FormatSource::MergeFormat
    } else if argument.eq_ignore_ascii_case("CHARFORMAT") {
        FormatSource::CharFormat
    } else {
        return false;
    };
    if *source != FormatSource::Default && *source != selected {
        // Neither ECMA-376 17.16.4.3.3 nor a Word control settles which of
        // the two wins; the first one is kept.
        crate::rejection::approximated(
            "Word field combines MERGEFORMAT and CHARFORMAT: the first applies",
            (),
        );
        return true;
    }
    *source = selected;
    true
}

/// The instruction as the shared renderer reads it, and the format source.
/// The renderer leaves a number in its default format for a switch it does
/// not interpret, as it does for a DOCX.
fn number_switches(instruction: &str) -> (String, FormatSource) {
    let instruction = instruction.trim();
    // Words with their byte offsets: only a format name changes.
    let mut words = instruction
        .split_whitespace()
        .map(|word| (word.as_ptr() as usize - instruction.as_ptr() as usize, word));
    words.next();
    let mut output = String::with_capacity(instruction.len());
    let mut copied = 0;
    let mut source = FormatSource::Default;
    let mut numeric = false;
    let mut interpreted = true;
    while let Some((_, switch)) = words.next() {
        if switch != "\\*" {
            interpreted = false;
            continue;
        }
        let Some((offset, argument)) = words.next() else {
            interpreted = false;
            break;
        };
        match canonical_number_format(argument) {
            Some(canonical) if !numeric => {
                numeric = true;
                output.push_str(&instruction[copied..offset]);
                output.push_str(canonical);
                copied = offset + argument.len();
            }
            _ => interpreted &= general_format(argument, &mut source),
        }
    }
    output.push_str(&instruction[copied..]);
    if !interpreted {
        crate::rejection::approximated(
            "Word page-number field switch without a renderer format",
            (),
        );
    }
    (output, source)
}

/// The format source of a DATE or TIME field. The shared renderer formats
/// the current date through the `\@` picture; without a picture, or with a
/// picture item it does not implement, it shows the stored result, as it
/// does for a DOCX.
fn date_switches(instruction: &str) -> FormatSource {
    let mut rest = instruction.trim_start();
    rest = rest[rest.find(char::is_whitespace).unwrap_or(rest.len())..].trim_start();
    let mut source = FormatSource::Default;
    let mut picture = false;
    while !rest.is_empty() {
        let (switch, tail) = rest.split_at(rest.find(char::is_whitespace).unwrap_or(rest.len()));
        let tail = tail.trim_start();
        if switch == "\\@" && tail.starts_with('"') {
            picture = true;
            rest = match tail[1..].find('"') {
                Some(close) => tail[close + 2..].trim_start(),
                None => "",
            };
            continue;
        }
        let (argument, tail) = tail.split_at(tail.find(char::is_whitespace).unwrap_or(tail.len()));
        if switch == "\\@" {
            picture = true;
        } else if switch == "\\*" {
            general_format(argument, &mut source);
        }
        rest = tail.trim_start();
    }
    if !picture {
        // Without a picture Word uses a language-dependent default format.
        crate::rejection::approximated(
            "Word date-time field without a picture shows its stored result",
            (),
        );
    }
    source
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::{tokenize_with_fields, Fields};

    fn with_types(text: &str, types: &[u8], flags: &[u8]) -> header_fields::Table {
        let mut entries = Vec::new();
        let mut cp = 0;
        let (mut begin, mut end) = (0, 0);
        for character in text.chars() {
            match character {
                '\u{13}' => {
                    entries.push((cp as u32, 0x13, types[begin]));
                    begin += 1;
                }
                '\u{14}' => entries.push((cp as u32, 0x14, 0)),
                '\u{15}' => {
                    entries.push((cp as u32, 0x15, flags[end]));
                    end += 1;
                }
                _ => {}
            }
            cp += character.len_utf16();
        }
        header_fields::Table::for_test(&entries, cp as u32)
    }

    fn tokens(text: &str, table: &header_fields::Table) -> Result<Vec<(Token, usize)>, String> {
        let fields = StoryFields::analyze(text, table, &[])?;
        let mut paragraphs = tokenize_with_fields(text, &mut Fields::default(), 0, true);
        fields.apply(0, &mut paragraphs)?;
        Ok(paragraphs
            .into_iter()
            .flat_map(|paragraph| paragraph.tokens)
            .collect())
    }

    fn page(instruction: &str, result: &str) -> Evaluated {
        Evaluated {
            field_type: "page",
            instruction: instruction.trim().into(),
            cached_result: result.into(),
            format_cp: 2,
            agreeing_result_cp: None,
            form_data_cp: None,
            drawn: None,
        }
    }

    #[test]
    fn evaluated_page_field_replaces_its_stored_result_in_place() {
        let text = "A\u{13}PAGE \\* roman\u{14}iv\u{15}B\r";
        let table = with_types(text, &[0x21], &[0x80]);
        let actual = tokens(text, &table).unwrap();
        let mut expected = page("PAGE \\* roman", "iv");
        expected.agreeing_result_cp = Some(16);
        assert_eq!(
            actual,
            vec![
                (Token::Text("A".into()), 0),
                (Token::EvaluatedField(Box::new(expected)), 1),
                (Token::Text("B".into()), 19),
            ]
        );
    }

    #[test]
    fn field_without_separator_is_still_evaluated_like_the_docx_parser() {
        // MS-DOC 2.8.25: the separator is present only when a result is
        // stored. The DOCX parser still computes PAGE at the end character.
        let text = "\u{13}PAGE  \u{15}\r";
        let table = with_types(text, &[0x21], &[0x00]);
        let actual = tokens(text, &table).unwrap();
        let mut expected = page("PAGE", "");
        expected.format_cp = 1;
        assert_eq!(actual, vec![(Token::EvaluatedField(Box::new(expected)), 0)]);
    }

    #[test]
    fn general_format_switch_selects_the_formatting_source() {
        let text =
            "\u{13} PAGE \\* MERGEFORMAT \u{14}7\u{15}\u{13} PAGE \\* CHARFORMAT \u{14}8\u{15}\r";
        let table = with_types(text, &[0x21, 0x21], &[0x80, 0x80]);
        let actual = tokens(text, &table).unwrap();
        let formats: Vec<_> = actual
            .iter()
            .map(|(token, _)| match token {
                Token::EvaluatedField(field) => (field.format_cp, field.agreeing_result_cp),
                _ => unreachable!(),
            })
            .collect();
        // MERGEFORMAT: first stored-result character. CHARFORMAT: the P.
        assert_eq!(formats, vec![(23, None), (27, None)]);
    }

    #[test]
    fn cached_fields_keep_results_and_nested_instructions_stay_hidden() {
        let text = "\u{13}HYPERLINK \\l \"x\"\u{14}\u{13}PAGEREF x \\h\u{14}3\u{15}\u{15}\r";
        let table = with_types(text, &[0x58, 0x25], &[0xc0, 0x80]);
        let actual = tokens(text, &table).unwrap();
        // The HYPERLINK field is the field form of <w:hyperlink>; its \l
        // location wins over the nested PAGEREF \h bookmark.
        let link = Link {
            href: None,
            anchor: Some("x".into()),
            in_toc: false,
        };
        assert_eq!(
            actual,
            vec![(
                Token::Linked(Box::new(Linked {
                    token: Token::Text("3".into()),
                    link: link.clone(),
                })),
                32
            )]
        );
        let toc = "\u{13}TOC \\h\u{14}\u{13}HYPERLINK \"http://a\" \\l \"y\"\u{14}T\u{15}U\u{15}\r";
        let table = with_types(toc, &[0x0d, 0x58], &[0xc0, 0x80]);
        let actual = tokens(toc, &table).unwrap();
        let link = Link {
            href: Some("http://a".into()),
            anchor: Some("y".into()),
            in_toc: true,
        };
        assert_eq!(
            actual
                .iter()
                .map(|(token, _)| token.clone())
                .collect::<Vec<_>>(),
            vec![
                Token::Linked(Box::new(Linked {
                    token: Token::Text("T".into()),
                    link,
                })),
                Token::Text("U".into()),
            ]
        );
        let hidden = "\u{13}IF \u{13}PAGE\u{14}1\u{15} = 1 \"a\"\u{14}a\u{15}\r";
        let table = with_types(hidden, &[0x07, 0x21], &[0xc0, 0x80]);
        let actual = tokens(hidden, &table).unwrap();
        assert_eq!(actual, vec![(Token::Text("a".into()), 21)]);
    }

    #[test]
    fn structure_must_match_the_field_table() {
        let text = "\u{13}PAGE\u{14}1\u{15}\r";
        // No table entries at all.
        assert!(tokens(text, &header_fields::Table::for_test(&[], 9))
            .unwrap_err()
            .contains("disagree"));
        // Listed type, but the flags claim no separator.
        let table = with_types(text, &[0x21], &[0x00]);
        assert!(tokens(text, &table).unwrap_err().contains("flags"));
        // Unbalanced characters.
        for text in ["\u{13}PAGE\r", "A\u{15}\r", "\u{14}\r"] {
            assert!(tokens(text, &header_fields::Table::for_test(&[], 9))
                .unwrap_err()
                .contains("unbalanced"));
        }
        // An unlisted XE field is allowed and hidden; other unlisted fields are not.
        let xe = "A\u{13} XE \"entry\" \u{15}B\r";
        assert_eq!(
            tokens(xe, &header_fields::Table::for_test(&[], 20)).unwrap(),
            vec![(Token::Text("A".into()), 0), (Token::Text("B".into()), 15)]
        );
        let other = "\u{13} REF x \u{15}\r";
        assert!(tokens(other, &header_fields::Table::for_test(&[], 20)).is_err());
        // A field type code that disagrees with an evaluated instruction.
        let wrong = with_types(text, &[0x1a], &[0x80]);
        assert!(tokens(text, &wrong).unwrap_err().contains("type disagrees"));
    }

    fn evaluated(text: &str, kind: u8, flags: u8) -> Vec<Evaluated> {
        let table = with_types(text, &[kind], &[flags]);
        tokens(text, &table)
            .unwrap()
            .into_iter()
            .filter_map(|(token, _)| match token {
                Token::EvaluatedField(field) => Some(*field),
                _ => None,
            })
            .collect()
    }

    fn best_effort() -> Vec<&'static str> {
        crate::rejection::take_best_effort()
            .into_iter()
            .map(|(_, site)| site)
            .collect()
    }

    #[test]
    fn fields_without_a_projection_stay_rejected() {
        for (text, kind, flags) in [
            ("\u{13} FORMCHECKBOX \u{14}\u{15}\r", 0x47, 0xa0),
            ("\u{13} FORMCHECKBOX \u{1}\u{14}x\u{15}\r", 0x47, 0xa0),
            ("\u{13} FORMCHECKBOX \u{1}\u{14}\u{15}\r", 0x47, 0xb0),
            ("\u{13}MACROBUTTON M Click\u{15}\r", 0x33, 0x00),
        ] {
            let table = with_types(text, &[kind], &[flags]);
            assert!(tokens(text, &table).is_err(), "{text:?}");
        }
    }

    #[test]
    fn locked_fields_display_their_stored_result() {
        // grffldEnd fLocked: Word does not update the field (Word control:
        // a locked DATE prints its stored result).
        let text = "\u{13}PAGE\u{14}7\u{15}\r";
        let table = with_types(text, &[0x21], &[0x90]);
        assert_eq!(
            tokens(text, &table).unwrap(),
            vec![(Token::Text("7".into()), 6)]
        );
    }

    #[test]
    fn display_and_edit_flags_do_not_stop_the_evaluation() {
        // fDiffer (0x01), fResultsDirty (0x04) and fResultsEdited (0x08):
        // Word prints the updated value (Word controls).
        best_effort();
        for flags in [0x81, 0x84, 0x88, 0x8c] {
            let fields = evaluated("\u{13}PAGE\u{14}99\u{15}\r", 0x21, flags);
            assert_eq!(fields.len(), 1, "{flags:02x}");
            assert_eq!(
                (fields[0].field_type, fields[0].cached_result.as_str()),
                ("page", "99")
            );
        }
        assert!(best_effort().is_empty());
    }

    #[test]
    fn private_results_are_not_displayed() {
        // MS-DOC 2.9.110 fPrivateResult: "not intended to be visible".
        let text = "A\u{13}REF x\u{14}sec\u{13}PAGE\u{14}1\u{15}ret\u{15}B\r";
        let table = with_types(text, &[0x03, 0x21], &[0xc0, 0xa0]);
        assert_eq!(
            tokens(text, &table).unwrap(),
            vec![(Token::Text("A".into()), 0), (Token::Text("B".into()), 23)]
        );
        // A private INCLUDEPICTURE result keeps its displayed picture.
        let picture = "\u{13}INCLUDEPICTURE \"x\"\u{14}\u{1}\u{15}\r";
        let table = with_types(picture, &[0x43], &[0xac]);
        assert_eq!(tokens(picture, &table).unwrap(), vec![(Token::Picture, 20)]);
    }

    #[test]
    fn page_number_switches_reach_the_renderer_in_its_spelling() {
        best_effort();
        // Word matches the format name without case and capitalizes like
        // its first letter (Word control: `\* ROMAN` prints `I`).
        for (instruction, projected) in [
            ("PAGE \\* ROMAN", "PAGE \\* Roman"),
            ("PAGE \\* roman", "PAGE \\* roman"),
            ("PAGE \\* Alphabetic", "PAGE \\* ALPHABETIC"),
            (
                "PAGE \\* alphabetic \\* mergeformat",
                "PAGE \\* alphabetic \\* mergeformat",
            ),
            ("PAGE \\* ARABICDASH", "PAGE \\* ArabicDash"),
        ] {
            let text = format!("\u{13}{instruction}\u{14}1\u{15}\r");
            let fields = evaluated(&text, 0x21, 0x80);
            assert_eq!(fields[0].instruction, projected);
        }
        assert!(best_effort().is_empty());
        // A switch without a renderer format leaves the default format.
        for instruction in [
            "PAGE \\# 0",
            "PAGE \\* CardText",
            "PAGE \\* roman \\* Arabic",
        ] {
            let text = format!("\u{13}{instruction}\u{14}1\u{15}\r");
            let fields = evaluated(&text, 0x21, 0x80);
            assert_eq!(fields[0].instruction, instruction);
            assert_eq!(
                best_effort(),
                ["Word page-number field switch without a renderer format"]
            );
        }
        // Both general formats: the first one selects the source.
        let fields = evaluated(
            "\u{13}PAGE \\* MERGEFORMAT \\* CHARFORMAT\u{14}1\u{15}\r",
            0x21,
            0x80,
        );
        assert_eq!(
            (fields[0].format_cp, fields[0].agreeing_result_cp),
            (35, None)
        );
        assert_eq!(
            best_effort(),
            ["Word field combines MERGEFORMAT and CHARFORMAT: the first applies"]
        );
    }

    #[test]
    fn evaluated_fields_replace_structured_results_and_nest_in_results() {
        // The stored result Word regenerates may hold anything.
        let text = "\u{13}PAGE\u{14}1\u{b}2\u{15}B\r";
        let table = with_types(text, &[0x21], &[0x80]);
        let actual = tokens(text, &table).unwrap();
        assert_eq!(actual.len(), 2);
        assert!(matches!(actual[0].0, Token::EvaluatedField(_)));
        assert_eq!(actual[1].0, Token::Text("B".into()));
        // A PAGE field inside a hyperlink result is evaluated (Word control).
        let nested = "\u{13}HYPERLINK \"http://a\"\u{14}\u{13}PAGE\u{14}1\u{15}\u{15}\r";
        let table = with_types(nested, &[0x58, 0x21], &[0xc0, 0x80]);
        let actual = tokens(nested, &table).unwrap();
        let [(Token::EvaluatedField(field), 22)] = actual.as_slice() else {
            panic!("{actual:?}");
        };
        assert_eq!(field.field_type, "page");
    }

    #[test]
    fn date_fields_leave_their_picture_to_the_renderer() {
        best_effort();
        // A picture item the renderer does not implement shows the stored
        // result there, as for a DOCX.
        let fields = evaluated("\u{13}DATE \\@ \"MMMM d\" \\l\u{14}x\u{15}\r", 0x1f, 0x80);
        assert_eq!(
            (fields[0].field_type, fields[0].instruction.as_str()),
            ("date", "DATE \\@ \"MMMM d\" \\l")
        );
        assert!(best_effort().is_empty());
        let fields = evaluated("\u{13}DATE\u{14}1/1/2000\u{15}\r", 0x1f, 0x80);
        assert_eq!(fields[0].cached_result, "1/1/2000");
        assert_eq!(
            best_effort(),
            ["Word date-time field without a picture shows its stored result"]
        );
    }

    #[test]
    fn an_evaluated_field_longer_than_the_scratch_is_a_resource_limit() {
        // The token replaces the whole field: an instruction or result the
        // scratch did not hold is not dropped.
        for text in [
            format!("\u{13}EQ {}\u{15}\r", "x".repeat(MAX_FIELD_TEXT + 8)),
            format!("\u{13}PAGE\u{14}{}\u{15}\r", "9".repeat(MAX_FIELD_TEXT + 8)),
        ] {
            let kind = if text.contains("EQ") { 0x31 } else { 0x21 };
            let flags = if text.contains('\u{14}') { 0x80 } else { 0x00 };
            let table = with_types(&text, &[kind], &[flags]);
            let error = tokens(&text, &table).unwrap_err();
            assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
        }
        // A cached field needs only its keyword.
        let text = format!(
            "\u{13}REF {}\u{14}r\u{15}\r",
            "x".repeat(MAX_FIELD_TEXT + 8)
        );
        let table = with_types(&text, &[0x03], &[0x80]);
        assert_eq!(tokens(&text, &table).unwrap().len(), 1);
    }

    #[test]
    fn a_field_replaces_only_the_fields_nested_in_it() {
        // Fields before an evaluated field stay; a private result drops the
        // evaluated field nested in it.
        let text = "\u{13}PAGE\u{14}1\u{15}\u{13}NUMPAGES\u{14}2\u{15}\u{13}REF x\u{14}\u{13}PAGE\u{14}3\u{15}\u{15}\r";
        let table = with_types(text, &[0x21, 0x1a, 0x03, 0x21], &[0x80, 0x80, 0xc0, 0xa0]);
        let kinds: Vec<_> = tokens(text, &table)
            .unwrap()
            .into_iter()
            .map(|(token, _)| match token {
                Token::EvaluatedField(field) => field.field_type,
                other => panic!("{other:?}"),
            })
            .collect();
        assert_eq!(kinds, ["page", "numPages"]);
    }

    #[test]
    fn symbol_fields_name_one_character() {
        let fields = evaluated("\u{13}SYMBOL 167 \\f \"Symbol\" \\s 14\u{15}\r", 0x39, 0x00);
        assert_eq!(fields[0].field_type, "symbol");
        assert_eq!(fields[0].format_cp, 1);
        assert_eq!(
            fields[0].drawn.as_deref(),
            Some(&Drawn::Symbol(SymbolForm {
                code: 167,
                font: Some("Symbol".into()),
                size_half_points: Some(28),
                unicode: false,
            }))
        );
        let fields = evaluated("\u{13}SYMBOL 0x2022 \\u\u{15}\r", 0x39, 0x00);
        assert_eq!(
            fields[0].drawn.as_deref(),
            Some(&Drawn::Symbol(SymbolForm {
                code: 0x2022,
                font: None,
                size_half_points: None,
                unicode: true,
            }))
        );
        // No character code: nothing is drawn.
        best_effort();
        assert!(evaluated("\u{13}SYMBOL \\f \"Symbol\"\u{15}\r", 0x39, 0x00).is_empty());
        assert_eq!(
            best_effort(),
            ["Word SYMBOL field without a character code"]
        );
    }

    #[test]
    fn hyperlinks_without_a_readable_target_link_nowhere() {
        best_effort();
        let text = "\u{13}HYPERLINK \"http://a\" \\z\u{14}T\u{15}\r";
        let table = with_types(text, &[0x58], &[0x80]);
        assert_eq!(
            tokens(text, &table).unwrap(),
            vec![(Token::Text("T".into()), 25)]
        );
        assert_eq!(
            best_effort(),
            ["Word hyperlink instruction without a readable target"]
        );
        // Nested hyperlinks: the innermost target.
        let text =
            "\u{13}HYPERLINK \"http://a\"\u{14}\u{13}HYPERLINK \"http://b\"\u{14}T\u{15}\u{15}\r";
        let table = with_types(text, &[0x58, 0x58], &[0xc0, 0x80]);
        let actual = tokens(text, &table).unwrap();
        let [(Token::Linked(linked), _)] = actual.as_slice() else {
            panic!("{actual:?}");
        };
        assert_eq!(linked.link.href.as_deref(), Some("http://b"));
        assert!(
            best_effort().contains(&"nested Word hyperlink fields link to the innermost target")
        );
    }

    #[test]
    fn check_boxes_carry_their_form_data_character() {
        let text = "A\u{13} FORMCHECKBOX \u{1}\u{14}\u{15}B\r";
        let table = with_types(text, &[0x47], &[0xa0]);
        let actual = tokens(text, &table).unwrap();
        let expected = Evaluated {
            field_type: "checkbox",
            instruction: "FORMCHECKBOX".into(),
            cached_result: String::new(),
            format_cp: 1,
            agreeing_result_cp: None,
            form_data_cp: Some(16),
            drawn: None,
        };
        assert_eq!(
            actual,
            vec![
                (Token::Text("A".into()), 0),
                (Token::EvaluatedField(Box::new(expected)), 1),
                (Token::Text("B".into()), 19),
            ]
        );
    }

    fn form_data(bits: u16, size: u16, name: &str, default: u16) -> Vec<u8> {
        let mut data = vec![0xff; 4];
        data.extend(bits.to_le_bytes());
        data.extend(0u16.to_le_bytes());
        data.extend(size.to_le_bytes());
        data.extend((name.encode_utf16().count() as u16).to_le_bytes());
        for unit in name.encode_utf16() {
            data.extend(unit.to_le_bytes());
        }
        data.extend(0u16.to_le_bytes());
        data.extend(default.to_le_bytes());
        data
    }

    #[test]
    fn check_box_state_follows_ffdata() {
        // iType 1, iRes in bits 2-6, iSize (automatic size) in bit 10.
        assert_eq!(
            checkbox_state(&form_data(1 | (1 << 2), 20, "Check1", 0)).unwrap(),
            (true, Some(10.0))
        );
        assert_eq!(
            checkbox_state(&form_data(1 | (1 << 10), 0, "", 1)).unwrap(),
            (false, None)
        );
        assert_eq!(
            checkbox_state(&form_data(1 | (25 << 2) | (1 << 10), 0, "", 0)).unwrap(),
            (false, None)
        );
        for data in [
            form_data(1 | (25 << 2), 20, "", 1),
            form_data(0, 20, "", 0),
            form_data(1 | (2 << 2), 20, "", 0),
            form_data(1, 1, "", 0),
            form_data(1, 20, "", 0)[..14].to_vec(),
        ] {
            assert!(checkbox_state(&data).is_err());
        }
    }

    #[test]
    fn eq_phonetic_guides_become_ruby_forms() {
        let text = "A\u{13}EQ \\* jc2 \\* \u{201c}Font:X\u{201d}  \\* hps16 \\o\\ad(\\s\\up 14(\u{3042}),\u{4e0a})\u{15}B\r";
        let table = with_types(text, &[0x31], &[0x00]);
        let actual = tokens(text, &table).unwrap();
        let Token::EvaluatedField(field) = &actual[1].0 else {
            panic!("ruby token: {actual:?}");
        };
        let Some(Drawn::Ruby(ruby)) = field.drawn.as_deref() else {
            panic!("ruby form: {field:?}");
        };
        assert_eq!(field.field_type, "ruby");
        assert_eq!(
            (ruby.guide.as_str(), ruby.base.as_str(), ruby.align),
            ("\u{3042}", "\u{4e0a}", "distributeSpace")
        );
        assert_eq!((ruby.guide_half_points, ruby.raise_pt), (Some(16), 14));
        let units: Vec<u16> = text.encode_utf16().collect();
        assert_eq!(units[ruby.guide_cp], 0x3042);
        assert_eq!(units[ruby.base_cp], 0x4e0a);
        assert_eq!(actual[2].0, Token::Text("B".into()));
    }

    #[test]
    fn other_eq_fields_become_equations() {
        let fields = evaluated("A\u{13}EQ \\f(1,2)\u{15}\r", 0x31, 0x00);
        assert_eq!(fields[0].field_type, "equation");
        // The first displayed character (the numerator) sizes the equation.
        assert_eq!(fields[0].format_cp, 8);
        let Some(Drawn::Equation(nodes)) = fields[0].drawn.as_deref() else {
            panic!("{fields:?}");
        };
        assert_eq!(
            serde_json::to_value(nodes).unwrap(),
            serde_json::json!([{
                "kind": "fraction",
                "num": [{ "kind": "run", "text": "1", "style": "roman" }],
                "den": [{ "kind": "run", "text": "2", "style": "roman" }],
            }])
        );
    }

    #[test]
    fn fields_may_not_cross_story_partitions() {
        let text = "\u{13}REF x\u{14}a\rb\u{15}\r";
        let table = with_types(text, &[0x03], &[0x80]);
        assert!(StoryFields::analyze(text, &table, &[]).is_ok());
        assert!(StoryFields::analyze(text, &table, &[9])
            .err()
            .unwrap()
            .contains("story boundary"));
    }
}
