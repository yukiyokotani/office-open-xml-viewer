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
//!   numbers follow pagination. They are projected onto the same `FieldRun`,
//!   limited to instructions whose switches the shared renderer interprets
//!   exactly; anything else is rejected instead of showing a stale value.
//! * Every other field displays its stored result, which stays ordinary runs,
//!   matching the DOCX parser's complex-field behavior. Fields never execute.
//! * Fields that Word draws without a stored result or as a control (EQ,
//!   MACROBUTTON, GOTOBUTTON, ADVANCE, SYMBOL without a result, form check boxes
//!   and drop-downs) are rejected until they are projected.
//!
//! A field result flagged fPrivateResult is displayed like any other result
//! only for INCLUDEPICTURE: the Word PDF of a DOC with such a picture result
//! shows the picture. MS-DOC only says the result "is not intended to be
//! visible to the user", so other private results with content stay rejected
//! until an Office control establishes their display.

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
}

/// Validated field structure of one aggregate story (main, header, footnote
/// or endnote document), with CPs relative to that document.
pub(super) struct StoryFields {
    evaluated: Vec<(usize, usize, Evaluated)>,
}

#[derive(Default)]
struct Open {
    begin: usize,
    parent: bool,
    listed_type: Option<u8>,
    hidden: bool,
    separator: Option<usize>,
    instruction: String,
    instruction_overflow: bool,
    instruction_controls: bool,
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
    pub(super) fn analyze(
        text: &str,
        table: &header_fields::Table,
        partitions: &[usize],
    ) -> Result<Self, String> {
        let mut stack: Vec<Open> = Vec::new();
        let mut evaluated = Vec::new();
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
                        return Err(unsupported("Word field nesting budget exceeded"));
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
                    if let Some(entry) = classify(&field, flags)? {
                        evaluated.push((field.begin, cp, entry));
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
        // Evaluated fields have no nested fields, so closing order is begin order.
        evaluated.sort_unstable_by_key(|(begin, ..)| *begin);
        Ok(Self { evaluated })
    }

    /// Replace each evaluated field's stored-result tokens in one already
    /// tokenized slice (`base_cp` is the slice's first CP) with one
    /// `Token::EvaluatedField`, placed at the field's begin character.
    pub(super) fn apply(&self, base_cp: usize, paragraphs: &mut [Paragraph]) -> Result<(), String> {
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
                if covered_until.is_some_and(|end| cp <= end) {
                    // Only stored-result text can lie inside an evaluated
                    // field; `classify` rejected every other content.
                    if !matches!(token, Token::Text(_)) {
                        return Err(unsupported("unexpected content in evaluated Word field"));
                    }
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
        }
        Ok(())
    }
}

impl Open {
    fn push(&mut self, character: char, cp: usize) {
        if self.separator.is_none() {
            if character < ' ' && character != '\t' {
                self.instruction_controls = true;
                return;
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

fn mismatch() -> String {
    unsupported("Word field characters disagree with the field table")
}

/// Field types whose field characters MS-DOC 2.8.25 forbids in a Plcfld. Word
/// stores them in the text only; none of them displays a result.
const UNLISTED_TYPES: [&str; 5] = ["XE", "TC", "RD", "TA", "PRIVATE"];

fn classify(field: &Open, flags: Option<u8>) -> Result<Option<Evaluated>, String> {
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
        return Ok(None);
    };
    // MS-DOC 2.9.110: fHasSep and fNested MUST describe the actual structure.
    if (flags & 0x80 != 0) != field.separator.is_some() || (flags & 0x40 != 0) != field.parent {
        return Err(unsupported("Word field flags disagree with its structure"));
    }
    if field.hidden {
        return Ok(None);
    }
    if flags & 0x01 != 0 {
        // fDiffer inverts the document's field-code display for this field.
        return Err(unsupported(
            "Word field displaying its instructions is not supported",
        ));
    }
    let (field_type, listed_type) = match keyword.as_str() {
        "PAGE" => ("page", 0x21),
        "NUMPAGES" => ("numPages", 0x1a),
        "DATE" => ("date", 0x1f),
        "TIME" => ("time", 0x20),
        "EQ" | "MACROBUTTON" | "GOTOBUTTON" | "ADVANCE" | "FORMCHECKBOX" | "FORMDROPDOWN" => {
            return Err(unsupported(format!(
                "Word {keyword} field display is not supported"
            )));
        }
        "SYMBOL" if field.separator.is_none() => {
            return Err(unsupported(
                "Word SYMBOL field without a stored result is not supported",
            ));
        }
        _ => {
            if flags & 0x20 != 0 && field.has_result_content() && keyword != "INCLUDEPICTURE" {
                return Err(unsupported(
                    "Word private field result display is not established",
                ));
            }
            return Ok(None);
        }
    };
    if field.listed_type != Some(listed_type) {
        return Err(unsupported(
            "Word evaluated field type disagrees with its instruction",
        ));
    }
    // fZombieEmbed, fResultsDirty, fResultsEdited, fLocked and fPrivateResult
    // each change whether Word re-evaluates or shows the stored result. None
    // has an Office control yet, so keep them closed for evaluated fields.
    if flags & 0x3e != 0 {
        return Err(unsupported(
            "Word evaluated field with edited, locked or private result is not supported",
        ));
    }
    if field.parent
        || field.child_in_instruction
        || field.child_in_result
        || field.instruction_controls
        || field.instruction_overflow
        || field.result_controls
        || field.result_overflow
    {
        return Err(unsupported(
            "Word evaluated field with nested or structural content is not supported",
        ));
    }
    let source = match field_type {
        "page" | "numPages" => number_switches(&field.instruction)?,
        _ => date_switches(&field.instruction)?,
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
        // instruction run. Word's choice is not specified, so a stored result
        // formatted differently is rejected by the projector.
        FormatSource::Default => (instruction_cp, field.result_cp),
    };
    Ok(Some(Evaluated {
        field_type,
        instruction: field.instruction.trim().to_string(),
        cached_result: field.result.clone(),
        format_cp,
        agreeing_result_cp,
    }))
}

#[derive(Debug, PartialEq, Eq)]
enum FormatSource {
    Default,
    CharFormat,
    MergeFormat,
}

/// ECMA-376 17.16.4.3.1 general numeric formats that the shared field renderer
/// maps one-to-one (case-sensitive arguments).
const NUMBER_FORMATS: [&str; 7] = [
    "Arabic",
    "ArabicDash",
    "Hex",
    "Roman",
    "roman",
    "ALPHABETIC",
    "alphabetic",
];

fn general_format(argument: &str, source: &mut FormatSource) -> Result<bool, String> {
    let selected = if argument.eq_ignore_ascii_case("MERGEFORMAT") {
        FormatSource::MergeFormat
    } else if argument.eq_ignore_ascii_case("CHARFORMAT") {
        FormatSource::CharFormat
    } else {
        return Ok(false);
    };
    if *source != FormatSource::Default && *source != selected {
        return Err(unsupported(
            "Word field combines MERGEFORMAT and CHARFORMAT",
        ));
    }
    *source = selected;
    Ok(true)
}

fn number_switches(instruction: &str) -> Result<FormatSource, String> {
    let mut words = instruction.split_whitespace().skip(1);
    let mut source = FormatSource::Default;
    let mut numeric = false;
    while let Some(switch) = words.next() {
        let argument = words.next();
        let supported = switch == "\\*"
            && argument.is_some_and(|argument| {
                if NUMBER_FORMATS.contains(&argument) && !numeric {
                    numeric = true;
                    return true;
                }
                general_format(argument, &mut source).unwrap_or(false)
            });
        if !supported {
            return Err(unsupported(format!(
                "unsupported Word page-number field instruction: {instruction}"
            )));
        }
    }
    Ok(source)
}

fn date_switches(instruction: &str) -> Result<FormatSource, String> {
    let mut rest = instruction.trim_start();
    rest = rest[rest.find(char::is_whitespace).unwrap_or(rest.len())..].trim_start();
    let mut source = FormatSource::Default;
    let mut picture = false;
    while !rest.is_empty() {
        let (switch, tail) = rest.split_at(rest.find(char::is_whitespace).unwrap_or(rest.len()));
        let tail = tail.trim_start();
        if switch == "\\@" && !picture && tail.starts_with('"') {
            let close = tail[1..]
                .find('"')
                .ok_or_else(|| unsupported("unterminated Word date-time picture"))?;
            validate_date_picture(&tail[1..1 + close])?;
            picture = true;
            rest = tail[close + 2..].trim_start();
            continue;
        }
        let (argument, tail) = tail.split_at(tail.find(char::is_whitespace).unwrap_or(tail.len()));
        if switch != "\\*" || !general_format(argument, &mut source)? {
            return Err(unsupported(format!(
                "unsupported Word date-time field instruction: {instruction}"
            )));
        }
        rest = tail.trim_start();
    }
    if !picture {
        // Without a picture Word uses a language-dependent default format;
        // the shared renderer then shows the stored result instead.
        return Err(unsupported(
            "Word date-time field without a picture is not supported",
        ));
    }
    Ok(source)
}

/// ECMA-376 17.16.4.1 picture items that the shared renderer formats without
/// locale data: numeric year, month, day, hour, minute and second items, the
/// AM/PM designators, quoted literals and punctuation. Month/weekday names and
/// calendar/era items are language dependent and stay rejected.
fn validate_date_picture(picture: &str) -> Result<(), String> {
    let characters: Vec<char> = picture.chars().collect();
    let reject = || unsupported(format!("unsupported Word date-time picture: {picture}"));
    let mut index = 0;
    while index < characters.len() {
        let character = characters[index];
        if character == '\'' {
            index += 1;
            loop {
                match characters.get(index) {
                    Some('\'') if characters.get(index + 1) == Some(&'\'') => index += 2,
                    Some('\'') => {
                        index += 1;
                        break;
                    }
                    Some(_) => index += 1,
                    None => return Err(reject()),
                }
            }
            continue;
        }
        if character.is_ascii_alphabetic() {
            let designator = characters[index..].iter().collect::<String>();
            if let Some(length) = ["AM/PM", "am/pm", "A/P", "a/p"]
                .iter()
                .find(|value| designator.starts_with(**value))
                .map(|value| value.len())
            {
                index += length;
                continue;
            }
            let run = characters[index..]
                .iter()
                .take_while(|value| **value == character)
                .count();
            let allowed = match character {
                'y' | 'Y' => run == 2 || run == 4,
                'M' | 'd' | 'D' | 'H' | 'h' | 'm' | 's' => run <= 2,
                _ => false,
            };
            if !allowed {
                return Err(reject());
            }
            index += run;
            continue;
        }
        if character.is_alphanumeric() || character == '"' || character == '\\' {
            return Err(reject());
        }
        index += 1;
    }
    Ok(())
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
        assert_eq!(actual.len(), 1);
        assert_eq!(actual[0].0, Token::Text("3".into()));
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

    #[test]
    fn fields_word_draws_itself_or_with_unverified_state_are_rejected() {
        for (text, kind, flags) in [
            ("\u{13}EQ \\o(a,b)\u{15}\r", 0x31, 0x00),
            ("\u{13} FORMCHECKBOX \u{14}\u{15}\r", 0x47, 0xa0),
            ("\u{13}MACROBUTTON M Click\u{15}\r", 0x33, 0x00),
            ("\u{13}SYMBOL 183\u{15}\r", 0x39, 0x00),
            ("\u{13}REF x\u{14}secret\u{15}\r", 0x03, 0xa0),
            ("\u{13}PAGE\u{14}1\u{15}\r", 0x21, 0x90),
            ("\u{13}PAGE\u{14}1\u{15}\r", 0x21, 0x81),
            ("\u{13}PAGE \\# 0\u{14}1\u{15}\r", 0x21, 0x80),
            ("\u{13}PAGE \\* ROMAN\u{14}I\u{15}\r", 0x21, 0x80),
            ("\u{13}PAGE \\* roman \\* Arabic\u{14}i\u{15}\r", 0x21, 0x80),
            ("\u{13}PAGE\u{14}1\u{b}2\u{15}\r", 0x21, 0x80),
            ("\u{13}DATE\u{14}1/1/2000\u{15}\r", 0x1f, 0x80),
            ("\u{13}DATE \\@ \"MMMM d\"\u{14}x\u{15}\r", 0x1f, 0x80),
            ("\u{13}DATE \\@ \"d\" \\l\u{14}x\u{15}\r", 0x1f, 0x80),
        ] {
            let table = with_types(text, &[kind], &[flags]);
            assert!(tokens(text, &table).is_err(), "{text:?}");
        }
        // Nested evaluated fields inside a cached result are rejected too.
        let nested = "\u{13}REF x\u{14}\u{13}PAGE\u{14}1\u{15}\u{15}\r";
        let table = with_types(nested, &[0x03, 0x21], &[0xc0, 0x80]);
        assert!(tokens(nested, &table).is_err());
        // A private INCLUDEPICTURE result keeps its displayed picture.
        let picture = "\u{13}INCLUDEPICTURE \"x\"\u{14}\u{1}\u{15}\r";
        let table = with_types(picture, &[0x43], &[0xac]);
        assert_eq!(tokens(picture, &table).unwrap(), vec![(Token::Picture, 20)]);
    }

    #[test]
    fn date_pictures_are_limited_to_locale_independent_items() {
        for picture in [
            "MM/dd/yyyy",
            "YYYY",
            "H:mm:ss",
            "h:mm AM/PM",
            "d 'of' M",
            "yy.M.d",
        ] {
            let instruction = format!("DATE \\@ \"{picture}\" \\* MERGEFORMAT");
            assert_eq!(
                date_switches(&instruction).unwrap(),
                FormatSource::MergeFormat,
                "{picture}"
            );
        }
        for picture in ["MMM", "dddd", "yyy", "ggg", "'open", "\\d"] {
            assert!(validate_date_picture(picture).is_err(), "{picture}");
        }
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
