//! OpenDocument text (ODT) → DOCX `Document`.
//!
//! The reader is fail-closed: every element, attribute and effective style
//! property it meets in rendered content is either mapped exactly, recognised
//! as inert, or rejects the whole document with a content-free reason code
//! (`crate::error::unsupported`). The admitted subset is deliberately narrow:
//! paragraphs and headings with spans of Latin-class text on a single page
//! style. Each widening comes with its own measurements against LibreOffice.

use crate::error::{invalid, unsupported};
use crate::package::{OdfPackage, SUPPORTED_VERSIONS};
use crate::script::{table_22, ScriptClass};
use crate::style::{Cascade, Family, Scope, StyleRef, StyleSheet, PARAGRAPH, TEXT};
use crate::xml::{
    self, NS_DRAW, NS_FO, NS_LOEXT, NS_OFFICE, NS_OFFICEOOO, NS_STYLE, NS_SVG, NS_TEXT, NS_XML,
};
use docx_model::{
    BodyElement, BreakType, DocParagraph, DocRun, Document, DocumentSettings, LineSpacing,
    RunFontFacts, SectionProps, TextRun,
};
use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap};

/// Media types an ODT source claims (Part 2 §3.3; templates share the model).
pub const MEDIA_TYPES: [&str; 2] = [
    "application/vnd.oasis.opendocument.text",
    "application/vnd.oasis.opendocument.text-template",
];

/// Implementation bound on nested inline elements walked recursively.
const MAX_INLINE_DEPTH: usize = 32;
/// Implementation bound on nested `text:section` elements.
const MAX_SECTION_DEPTH: usize = 32;
/// Implementation bound on retained text, the dominant construction cost of
/// this reader; the DOCX cursor additionally enforces its JSON ceilings.
const MAX_RETAINED_TEXT_BYTES: usize = 64 * 1024 * 1024;
/// Implementation bound on paragraphs.
const MAX_PARAGRAPHS: usize = 1_000_000;

pub fn read(package: &OdfPackage) -> Result<Document, String> {
    let styles_text = package.read_text("styles.xml")?;
    let content_text = package.read_text("content.xml")?;
    let styles_doc = xml::parse(&styles_text)?;
    let content_doc = xml::parse(&content_text)?;
    let styles_root = styles_doc.root_element();
    let content_root = content_doc.root_element();
    if !xml::is(styles_root, NS_OFFICE, "document-styles")
        || !xml::is(content_root, NS_OFFICE, "document-content")
    {
        return Err(invalid("styles.xml or content.xml has an unexpected root"));
    }
    // Part 3 §19.388: one declared version for the document's parts.
    let version = content_root.attribute((NS_OFFICE, "version"));
    if !version.is_some_and(|v| SUPPORTED_VERSIONS.contains(&v))
        || styles_root.attribute((NS_OFFICE, "version")) != version
    {
        return Err(unsupported("odf-version"));
    }

    let mut sheet = StyleSheet::default();
    let mut fonts = FontFaces::default();
    for (root, scope) in [(styles_root, Scope::Styles), (content_root, Scope::Content)] {
        for child in root.children().filter(|n| n.is_element()) {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_OFFICE), "font-face-decls") => fonts.add(child)?,
                (Some(NS_OFFICE), "styles") => sheet.add_common(child)?,
                (Some(NS_OFFICE), "automatic-styles") => sheet.add_automatic(scope, child)?,
                (Some(NS_OFFICE), "master-styles" | "body") => {}
                (Some(NS_OFFICE), "scripts") if child.children().all(|n| !n.is_element()) => {}
                _ => return Err(unsupported("document-part")),
            }
        }
    }
    check_document_wide_configuration(styles_root)?;
    let page = PageStyle::read(styles_root)?;

    let body = content_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|b| b.children().find(|n| n.is_element()))
        .filter(|n| xml::is(*n, NS_OFFICE, "text"))
        .ok_or_else(|| unsupported("not-a-text-document"))?;

    let mut reader = Reader {
        sheet: &sheet,
        fonts: &fonts,
        used_faces: RefCell::new(Vec::new()),
        page_master: page.name.as_str(),
        paragraphs: Vec::new(),
        retained_text_bytes: 0,
    };
    reader.walk_blocks(body, 0)?;
    let used = used_font_table(&reader.used_faces.borrow())?;
    let body = reader.finish()?;

    let default_tab_stop = sheet
        .default_style(Family::Paragraph)
        .and_then(|style| {
            Cascade {
                styles: vec![style],
                group: PARAGRAPH,
            }
            .get(NS_STYLE, "tab-stop-distance")
        })
        .map(length_pt)
        .transpose()?;

    Ok(Document {
        section: page.section,
        body,
        settings: Some(DocumentSettings {
            default_tab_stop,
            ..DocumentSettings::default()
        }),
        font_family_classes: metadata(&used, |face| face.generic),
        font_family_pitches: metadata(&used, |face| face.pitch),
        font_family_charsets: metadata(&used, |face| face.charset),
        ..Document::default()
    })
}

/// `style:font-face` declarations (Part 3 §16.23): the face a style
/// references, its family and the metadata the DOCX font table carries.
#[derive(Default)]
struct FontFaces {
    faces: HashMap<String, Face>,
}

#[derive(Clone, PartialEq)]
struct Face {
    family: String,
    generic: Option<&'static str>,
    pitch: Option<&'static str>,
    charset: Option<&'static str>,
}

impl FontFaces {
    fn add(&mut self, decls: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for face in decls.children().filter(|n| n.is_element()) {
            if !xml::is(face, NS_STYLE, "font-face") {
                return Err(unsupported("font-face"));
            }
            if face.children().any(|n| n.is_element()) {
                // svg:font-face-src: embedded font files.
                return Err(unsupported("embedded-font"));
            }
            let name = face
                .attribute((NS_STYLE, "name"))
                .ok_or_else(|| invalid("a font face has no name"))?;
            let family = face
                .attribute((NS_SVG, "font-family"))
                .ok_or_else(|| invalid("a font face has no family"))?;
            let value = Face {
                family: unquote_family(family)?,
                generic: generic_class(face.attribute((NS_STYLE, "font-family-generic")))?,
                pitch: pitch(face.attribute((NS_STYLE, "font-pitch")))?,
                charset: charset(face.attribute((NS_STYLE, "font-charset")))?,
            };
            match self.faces.get(name) {
                Some(existing) if *existing != value => {
                    return Err(unsupported("font-face-conflict"))
                }
                _ => {
                    self.faces.insert(name.to_string(), value);
                }
            }
        }
        Ok(())
    }
}

/// The DOCX font table (classes, pitches, charsets) is keyed by family name.
/// Producers may declare one family under several faces with different
/// metadata (LibreOffice does for fallback faces); only the faces admitted
/// text uses must agree.
fn used_font_table(used: &[Face]) -> Result<HashMap<String, Face>, String> {
    let mut table: HashMap<String, Face> = HashMap::new();
    for face in used {
        if let Some(existing) = table.insert(face.family.clone(), face.clone()) {
            if existing != *face {
                return Err(unsupported("font-face-conflict"));
            }
        }
    }
    Ok(table)
}

fn metadata(
    table: &HashMap<String, Face>,
    field: impl Fn(&Face) -> Option<&'static str>,
) -> BTreeMap<String, String> {
    table
        .values()
        .filter_map(|face| field(face).map(|value| (face.family.clone(), value.to_string())))
        .collect()
}

/// ODF `style:font-family-generic` (§20.278) and DOCX `w:family` (ECMA-376
/// §17.8.3.10) share their vocabulary; `system` has no DOCX class.
fn generic_class(value: Option<&str>) -> Result<Option<&'static str>, String> {
    Ok(match value {
        None | Some("system") => None,
        Some("roman") => Some("roman"),
        Some("swiss") => Some("swiss"),
        Some("modern") => Some("modern"),
        Some("decorative") => Some("decorative"),
        Some("script") => Some("script"),
        Some(_) => return Err(invalid("unknown font family generic class")),
    })
}

fn pitch(value: Option<&str>) -> Result<Option<&'static str>, String> {
    Ok(match value {
        None => None,
        Some("fixed") => Some("fixed"),
        Some("variable") => Some("variable"),
        Some(_) => return Err(invalid("unknown font pitch")),
    })
}

/// §20.268 `style:font-charset` is an encoding name. A Unicode encoding (or
/// none) needs no DOCX charset; `x-symbol` is the symbol charset (02).
fn charset(value: Option<&str>) -> Result<Option<&'static str>, String> {
    Ok(match value {
        None => None,
        Some(v) if v.eq_ignore_ascii_case("utf-8") || v.eq_ignore_ascii_case("utf-16") => None,
        Some("x-symbol") => Some("02"),
        Some(_) => return Err(unsupported("font-charset")),
    })
}

/// `svg:font-family` is a CSS font-family value; the reader admits exactly one
/// family name, quoted or not.
fn unquote_family(value: &str) -> Result<String, String> {
    let trimmed = value.trim();
    let name = trimmed
        .strip_prefix('\'')
        .and_then(|v| v.strip_suffix('\''))
        .or_else(|| trimmed.strip_prefix('"').and_then(|v| v.strip_suffix('"')))
        .unwrap_or(trimmed);
    if name.is_empty() || name.contains(',') {
        return Err(unsupported("font-family-list"));
    }
    Ok(name.to_string())
}

/// Document-wide configuration in `office:styles` that changes rendering.
fn check_document_wide_configuration(styles_root: roxmltree::Node<'_, '_>) -> Result<(), String> {
    let Some(common) = styles_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "styles"))
    else {
        return Ok(());
    };
    for node in common.children().filter(|n| n.is_element()) {
        if xml::is(node, NS_TEXT, "linenumbering-configuration")
            && node.attribute((NS_TEXT, "number-lines")) == Some("true")
        {
            return Err(unsupported("line-numbering"));
        }
        if xml::is(node, NS_TEXT, "outline-style") {
            // Headings are numbered by the outline style (§16.34); admitted
            // only when no level produces a label.
            for level in node.children().filter(|n| n.is_element()) {
                let labelled = ["num-format", "num-prefix", "num-suffix"]
                    .iter()
                    .any(|name| {
                        level
                            .attribute((NS_STYLE, *name))
                            .is_some_and(|v| !v.is_empty())
                    });
                if labelled {
                    return Err(unsupported("heading-numbering"));
                }
            }
        }
    }
    Ok(())
}

/// The single page style admitted documents use.
struct PageStyle {
    name: String,
    section: SectionProps,
}

impl PageStyle {
    /// The first master page is the initial page style (LibreOffice writes its
    /// default page style first); a paragraph may name it again. Admitted: one
    /// master page with no displayed header/footer and no successor.
    fn read(styles_root: roxmltree::Node<'_, '_>) -> Result<Self, String> {
        let master = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "master-styles"))
            .and_then(|m| m.children().find(|n| xml::is(*n, NS_STYLE, "master-page")))
            .ok_or_else(|| invalid("styles.xml has no master page"))?;
        let name = master
            .attribute((NS_STYLE, "name"))
            .ok_or_else(|| invalid("a master page has no name"))?
            .to_string();
        if master
            .attribute((NS_STYLE, "next-style-name"))
            .is_some_and(|next| next != name)
        {
            return Err(unsupported("master-page-succession"));
        }
        if master
            .children()
            .filter(|n| n.is_element())
            .any(|region| region.attribute((NS_STYLE, "display")) != Some("false"))
        {
            return Err(unsupported("header-footer"));
        }
        let layout_name = master
            .attribute((NS_STYLE, "page-layout-name"))
            .ok_or_else(|| invalid("a master page has no page layout"))?;
        let layout = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "automatic-styles"))
            .and_then(|a| {
                a.children().find(|n| {
                    xml::is(*n, NS_STYLE, "page-layout")
                        && n.attribute((NS_STYLE, "name")) == Some(layout_name)
                })
            })
            .ok_or_else(|| invalid("a master page names a missing page layout"))?;
        if layout
            .attribute((NS_STYLE, "page-usage"))
            .is_some_and(|usage| usage != "all")
        {
            return Err(unsupported("page-usage"));
        }
        let mut section = SectionProps::default();
        let mut saw_properties = false;
        for child in layout.children().filter(|n| n.is_element()) {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_STYLE), "page-layout-properties") => {
                    read_page_properties(child, &mut section)?;
                    saw_properties = true;
                }
                // Area styles of header/footer regions, which are not displayed.
                (Some(NS_STYLE), "header-style" | "footer-style") => {}
                _ => return Err(unsupported("page-layout-child")),
            }
        }
        if !saw_properties {
            return Err(invalid("a page layout has no properties"));
        }
        Ok(Self { name, section })
    }
}

fn read_page_properties(
    node: roxmltree::Node<'_, '_>,
    section: &mut SectionProps,
) -> Result<(), String> {
    let mut width = None;
    let mut height = None;
    for attribute in node.attributes() {
        let value = attribute.value();
        match (attribute.namespace().unwrap_or(""), attribute.name()) {
            (NS_FO, "page-width") => width = Some(length_pt(value)?),
            (NS_FO, "page-height") => height = Some(length_pt(value)?),
            (NS_FO, "margin-top") => section.margin_top = length_pt(value)?,
            (NS_FO, "margin-bottom") => section.margin_bottom = length_pt(value)?,
            (NS_FO, "margin-left") => section.margin_left = length_pt(value)?,
            (NS_FO, "margin-right") => section.margin_right = length_pt(value)?,
            // Numbering format, printing and spreadsheet attributes are inert
            // for a text document without page fields.
            (
                NS_STYLE,
                "num-format"
                | "num-letter-sync"
                | "print-orientation"
                | "print"
                | "print-page-order"
                | "first-page-number"
                | "scale-to"
                | "scale-to-pages"
                | "table-centering"
                | "footnote-max-height"
                | "paper-tray-name"
                | "register-truth-ref-style-name",
            ) => {}
            (NS_STYLE, "writing-mode") if matches!(value, "lr-tb" | "page") => {}
            (NS_STYLE, "layout-grid-mode") if value == "none" => {}
            (NS_STYLE, name) if name.starts_with("layout-grid-") => {}
            (NS_LOEXT, "margin-gutter") if length_pt(value)? == 0.0 => {}
            (NS_STYLE, "shadow") if value == "none" => {}
            (NS_FO, "background-color") if value == "transparent" => {}
            (NS_DRAW, "fill") if value == "none" => {}
            _ => return Err(unsupported("page-layout-property")),
        }
    }
    for child in node.children().filter(|n| n.is_element()) {
        match (child.tag_name().namespace(), child.tag_name().name()) {
            // The footnote separator is inert without notes (rejected).
            (Some(NS_STYLE), "footnote-sep") => {}
            (Some(NS_STYLE), "background-image")
                if child.children().all(|n| !n.is_element()) && child.attributes().len() == 0 => {}
            (Some(NS_STYLE), "columns")
                if child.attribute((NS_FO, "column-count")).unwrap_or("1") == "1"
                    && child.children().all(|n| !n.is_element()) => {}
            _ => return Err(unsupported("page-layout-property")),
        }
    }
    section.page_width = width.ok_or_else(|| invalid("a page layout has no width"))?;
    section.page_height = height.ok_or_else(|| invalid("a page layout has no height"))?;
    Ok(())
}

/// One collected paragraph before inter-paragraph spacing is resolved.
struct Collected {
    paragraph: DocParagraph,
    top: f64,
    bottom: f64,
    contextual: bool,
    /// The named (common) style compared by contextual spacing.
    named_style: Option<String>,
    page_break_after: bool,
}

struct Reader<'a> {
    sheet: &'a StyleSheet,
    fonts: &'a FontFaces,
    used_faces: RefCell<Vec<Face>>,
    page_master: &'a str,
    paragraphs: Vec<Collected>,
    retained_text_bytes: usize,
}

#[derive(Clone)]
struct TextContext<'a> {
    styles: Vec<StyleRef<'a>>,
}

enum Token<'a> {
    Char(char, TextContext<'a>),
    /// `text:s`: one space that is never collapsed.
    Space(TextContext<'a>),
    Tab(TextContext<'a>),
    LineBreak,
}

impl<'a> Reader<'a> {
    fn walk_blocks(&mut self, parent: roxmltree::Node<'_, '_>, depth: usize) -> Result<(), String> {
        if depth > MAX_SECTION_DEPTH {
            return Err(unsupported("section-depth"));
        }
        for node in parent.children() {
            if node.is_text() {
                if node.text().is_some_and(|t| !t.trim().is_empty()) {
                    return Err(invalid("text outside a paragraph"));
                }
                continue;
            }
            if !node.is_element() {
                continue;
            }
            match (node.tag_name().namespace(), node.tag_name().name()) {
                (Some(NS_TEXT), "p") => self.paragraph(node, None)?,
                (Some(NS_TEXT), "h") => {
                    let level = node
                        .attribute((NS_TEXT, "outline-level"))
                        .map(|v| v.parse::<u32>().map_err(|_| invalid("bad outline level")))
                        .transpose()?;
                    self.paragraph(node, level)?
                }
                // A producer hint (§5.6) and declarations without rendering.
                (
                    Some(NS_TEXT),
                    "soft-page-break"
                    | "sequence-decls"
                    | "variable-decls"
                    | "user-field-decls"
                    | "dde-connection-decls",
                ) => {}
                (Some(NS_OFFICE), "forms") if node.children().all(|n| !n.is_element()) => {}
                (Some(NS_TEXT), "tracked-changes") if node.children().all(|n| !n.is_element()) => {}
                (Some(NS_TEXT), "section") => {
                    // A plain section only groups paragraphs; section styles
                    // (columns, backgrounds), conditions and hidden sections
                    // change rendering.
                    for attribute in node.attributes() {
                        match (attribute.namespace(), attribute.name()) {
                            (Some(NS_TEXT), "name") | (Some(NS_XML), "id") => {}
                            (Some(NS_TEXT), "display") if attribute.value() == "true" => {}
                            _ => return Err(unsupported("section")),
                        }
                    }
                    self.walk_blocks(node, depth + 1)?
                }
                _ => return Err(unsupported("body-element")),
            }
        }
        Ok(())
    }

    fn paragraph(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        outline_level: Option<u32>,
    ) -> Result<(), String> {
        if self.paragraphs.len() >= MAX_PARAGRAPHS {
            return Err("OpenDocument text exceeds the paragraph limit".into());
        }
        for attribute in node.attributes() {
            match (attribute.namespace(), attribute.name()) {
                (
                    Some(NS_TEXT),
                    "style-name" | "outline-level" | "id" | "is-list-header" | "restart-numbering",
                ) => {}
                (Some(NS_XML), "id") => {}
                _ => return Err(unsupported("paragraph-attribute")),
            }
        }
        let mut chain: Vec<StyleRef<'a>> = match node.attribute((NS_TEXT, "style-name")) {
            Some(name) => self.sheet.chain(
                self.sheet
                    .find(Scope::Content, Family::Paragraph, name)
                    .ok_or_else(|| invalid("a paragraph names a missing style"))?,
            )?,
            None => Vec::new(),
        };
        let named_style = chain
            .iter()
            .find(|s| !s.style().automatic)
            .map(|s| s.style().name.clone());
        if let Some(default) = self.sheet.default_style(Family::Paragraph) {
            chain.push(default);
        }
        let properties = Cascade {
            styles: chain.clone(),
            group: PARAGRAPH,
        };

        // §19.501: naming the page style on a later paragraph starts a page.
        let master = chain.iter().find_map(|s| {
            s.style()
                .attributes
                .get(&(NS_STYLE.to_string(), "master-page-name".to_string()))
                .map(String::as_str)
        });
        let mut page_break_before = false;
        if let Some(master) = master.filter(|name| !name.is_empty()) {
            if master != self.page_master {
                return Err(unsupported("master-page-change"));
            }
            page_break_before = !self.paragraphs.is_empty();
            if properties
                .get(NS_STYLE, "page-number")
                .is_some_and(|v| v != "auto")
            {
                return Err(unsupported("page-number-restart"));
            }
        }

        let mut collected = paragraph_properties(&properties)?;
        collected.named_style = named_style;
        collected.paragraph.page_break_before |= page_break_before;
        collected.paragraph.outline_level = outline_level.map(|level| level.saturating_sub(1));

        // The paragraph's own text properties describe its mark (empty
        // paragraph height) and the text outside spans.
        let context = TextContext { styles: chain };
        let mark = self.run_properties(&context, ScriptClass::Latin)?;
        collected.paragraph.paragraph_mark_font_facts = Some(mark.font_facts());
        collected.paragraph.default_font_size = Some(mark.size);
        collected.paragraph.default_font_family = Some(mark.family.clone());

        let mut tokens = Vec::new();
        self.inline(node, &context, &mut tokens, 0)?;
        if tokens.iter().any(|t| matches!(t, Token::Tab(_)))
            && (collected.paragraph.indent_left != 0.0 || collected.paragraph.indent_first != 0.0)
        {
            // LibreOffice measures tab positions from the paragraph indent
            // unless TabsRelativeToIndent is off; DOCX from the margin. They
            // coincide only without an indent.
            return Err(unsupported("tab-with-indent"));
        }
        self.emit_runs(collapse(tokens), &mut collected.paragraph)?;
        self.paragraphs.push(collected);
        Ok(())
    }

    /// Inline content (§6.1): character data, spans and the white-space
    /// elements, flattened into tokens for §6.1.2 collapsing.
    fn inline(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        context: &TextContext<'a>,
        tokens: &mut Vec<Token<'a>>,
        depth: usize,
    ) -> Result<(), String> {
        if depth > MAX_INLINE_DEPTH {
            return Err(unsupported("inline-depth"));
        }
        for child in node.children() {
            if child.is_text() {
                let text = child.text().unwrap_or("");
                self.retained_text_bytes = self
                    .retained_text_bytes
                    .checked_add(text.len())
                    .filter(|total| *total <= MAX_RETAINED_TEXT_BYTES)
                    .ok_or_else(|| {
                        "OpenDocument text exceeds the retained text limit".to_string()
                    })?;
                tokens.extend(text.chars().map(|c| Token::Char(c, context.clone())));
                continue;
            }
            if !child.is_element() {
                continue;
            }
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "span") => {
                    if child
                        .attributes()
                        .any(|a| (a.namespace(), a.name()) != (Some(NS_TEXT), "style-name"))
                    {
                        return Err(unsupported("span-attribute"));
                    }
                    let nested = match child.attribute((NS_TEXT, "style-name")) {
                        // §16.2: a text style continues in the enclosing text
                        // styles, then the paragraph's style and its default.
                        Some(name) => {
                            let style = self
                                .sheet
                                .find(Scope::Content, Family::Text, name)
                                .ok_or_else(|| invalid("a span names a missing style"))?;
                            let mut styles = self.sheet.chain(style)?;
                            styles.extend(context.styles.iter().copied());
                            TextContext { styles }
                        }
                        None => context.clone(),
                    };
                    self.inline(child, &nested, tokens, depth + 1)?;
                }
                (Some(NS_TEXT), "s") => {
                    let count = child
                        .attribute((NS_TEXT, "c"))
                        .map(|v| v.parse::<usize>().map_err(|_| invalid("bad text:s count")))
                        .transpose()?
                        .unwrap_or(1);
                    if count == 0 || count > 4096 {
                        return Err(invalid("text:s count is out of range"));
                    }
                    tokens.extend((0..count).map(|_| Token::Space(context.clone())));
                }
                (Some(NS_TEXT), "tab") => tokens.push(Token::Tab(context.clone())),
                (Some(NS_TEXT), "line-break") => tokens.push(Token::LineBreak),
                (
                    Some(NS_TEXT),
                    "soft-page-break"
                    | "bookmark"
                    | "bookmark-start"
                    | "bookmark-end"
                    | "reference-mark"
                    | "reference-mark-start"
                    | "reference-mark-end",
                ) => {}
                _ => return Err(unsupported("inline-element")),
            }
        }
        Ok(())
    }

    /// Script classes (§20.358 Table 22). A character Table 22 does not map
    /// takes the class of the preceding character — observed LibreOffice
    /// behaviour for U+0020, U+00A0, U+2014 and U+1F600 between and after
    /// Latin and Asian text (see the measurement notes in the design; pending
    /// user approval as a compatibility rule). With no preceding character the
    /// class is admitted only when every class resolves to the same face.
    fn emit_runs(
        &mut self,
        tokens: Vec<Token<'a>>,
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        let mut previous: Option<ScriptClass> = None;
        for token in tokens {
            let (c, context) = match token {
                Token::LineBreak => {
                    paragraph.runs.push(DocRun::Break {
                        break_type: BreakType::Line,
                    });
                    continue;
                }
                Token::Char(c, context) => (c, context),
                Token::Space(context) => (' ', context),
                Token::Tab(context) => ('\t', context),
            };
            let class = match table_22(c) {
                Some(class) => class,
                None => match previous {
                    Some(class) => class,
                    None => self.uniform_class(&context)?,
                },
            };
            if class != ScriptClass::Latin {
                // Asian and complex text (line breaking, punctuation, autospace
                // and shaping rules) is not yet measured against LibreOffice.
                return Err(unsupported("script-class"));
            }
            previous = Some(class);
            let properties = self.run_properties(&context, class)?;
            if let Some(DocRun::Text(last)) = paragraph.runs.last_mut() {
                if properties.continues(last) {
                    last.text.push(c);
                    continue;
                }
            }
            let mut run = properties.text_run();
            run.text.push(c);
            paragraph.runs.push(DocRun::Text(Box::new(run)));
        }
        Ok(())
    }

    fn uniform_class(&self, context: &TextContext<'a>) -> Result<ScriptClass, String> {
        let latin = self.run_properties(context, ScriptClass::Latin)?;
        let asian = self.run_properties(context, ScriptClass::Asian)?;
        let complex = self.run_properties(context, ScriptClass::Complex)?;
        if latin == asian && latin == complex {
            Ok(ScriptClass::Latin)
        } else {
            Err(unsupported("script-initial-weak-character"))
        }
    }

    fn run_properties(
        &self,
        context: &TextContext<'a>,
        class: ScriptClass,
    ) -> Result<RunProperties, String> {
        let cascade = Cascade {
            styles: context.styles.clone(),
            group: TEXT,
        };
        let (properties, face) = RunProperties::resolve(&cascade, class, self.fonts)?;
        let mut used = self.used_faces.borrow_mut();
        if !used.contains(&face) {
            used.push(face);
        }
        Ok(properties)
    }

    /// Paragraph gaps (measured against LibreOffice): the previous paragraph's
    /// bottom margin and the next one's top margin add; `style:contextual-
    /// spacing` on both paragraphs of one named style removes the gap.
    /// LibreOffice decides by the previous paragraph's flag alone, which
    /// differs from §20.255 ("both") when only one paragraph sets it, so that
    /// mixed case is rejected. The gap is emitted as the next paragraph's
    /// space before with space after 0, so the DOCX max(after, before) rule
    /// reproduces it.
    fn finish(self) -> Result<Vec<BodyElement>, String> {
        let mut out = Vec::with_capacity(self.paragraphs.len());
        let mut previous: Option<&Collected> = None;
        let mut break_pending = false;
        for collected in &self.paragraphs {
            let mut paragraph = collected.paragraph.clone();
            paragraph.space_before = match previous {
                None => collected.top,
                Some(prev) => {
                    let same =
                        prev.named_style.is_some() && prev.named_style == collected.named_style;
                    if same && prev.contextual != collected.contextual {
                        return Err(unsupported("contextual-spacing-mixed"));
                    }
                    if same && prev.contextual {
                        0.0
                    } else {
                        prev.bottom + collected.top
                    }
                }
            };
            paragraph.space_after = 0.0;
            paragraph.contextual_spacing = false;
            paragraph.page_break_before |= break_pending;
            break_pending = collected.page_break_after;
            out.push(BodyElement::Paragraph(Box::new(paragraph)));
            previous = Some(collected);
        }
        Ok(out)
    }
}

/// §6.1.2 white-space collapsing over a paragraph's character data: tab, CR
/// and LF become spaces, leading and trailing spaces are removed and runs of
/// spaces collapse to one. `text:s`, `text:tab` and `text:line-break` keep
/// the characters they represent and are not collapsed.
fn collapse(tokens: Vec<Token<'_>>) -> Vec<Token<'_>> {
    let mut out: Vec<Token<'_>> = Vec::with_capacity(tokens.len());
    let mut pending: Option<Token<'_>> = None;
    let mut started = false;
    for token in tokens {
        match token {
            Token::Char(' ' | '\t' | '\r' | '\n', context) => {
                if started && pending.is_none() {
                    pending = Some(Token::Char(' ', context));
                }
            }
            Token::Char(c, context) => {
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                started = true;
                out.push(Token::Char(c, context));
            }
            element => {
                // Spaces adjacent to an element are character data on either
                // side of it; only consecutive character-data spaces merge.
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                started = true;
                out.push(element);
            }
        }
    }
    // A trailing collapsed space is removed (step 5).
    out
}

/// Effective text properties of one script class.
#[derive(Clone, PartialEq, Debug)]
struct RunProperties {
    family: String,
    size: f64,
    bold: bool,
    italic: bool,
    underline: Option<&'static str>,
    strike: Option<&'static str>,
    color: Option<String>,
    color_auto: bool,
    background: Option<String>,
    letter_spacing: Option<f64>,
    scale: Option<f64>,
    position: Option<f64>,
    caps: bool,
    kerning: Option<f64>,
    lang_east_asia: Option<String>,
    lang_bidi: Option<String>,
}

impl RunProperties {
    fn resolve(
        cascade: &Cascade<'_>,
        class: ScriptClass,
        fonts: &FontFaces,
    ) -> Result<(Self, Face), String> {
        check_text_keys(cascade)?;
        let suffix = match class {
            ScriptClass::Latin => "",
            ScriptClass::Asian => "-asian",
            ScriptClass::Complex => "-complex",
        };
        // Western properties are fo: attributes; the script variants are
        // style: attributes with a -asian / -complex suffix (§20.190 et al.).
        let scripted = |name: &str| -> Option<&str> {
            if suffix.is_empty() {
                cascade.get(NS_FO, name)
            } else {
                cascade.get(NS_STYLE, &format!("{name}{suffix}"))
            }
        };
        let face = resolve_family(cascade, suffix, fonts)?;
        let family = face.family.clone();
        let size = match scripted("font-size") {
            Some(value) if value.ends_with('%') => return Err(unsupported("relative-font-size")),
            Some(value) => length_pt(value)?,
            // ODF defines no initial font size; LibreOffice's default styles
            // always set one.
            None => return Err(unsupported("font-size-unresolved")),
        };
        let bold = match scripted("font-weight") {
            None | Some("normal") | Some("400") => false,
            Some("bold") | Some("700") => true,
            Some(_) => return Err(unsupported("font-weight")),
        };
        let italic = match scripted("font-style") {
            None | Some("normal") => false,
            Some("italic") => true,
            Some(_) => return Err(unsupported("font-style")),
        };
        let line = |prefix: &str| -> Result<Option<&'static str>, String> {
            let style = cascade
                .get(NS_STYLE, &format!("{prefix}-style"))
                .unwrap_or("none");
            if style == "none" {
                return Ok(None);
            }
            let kind = cascade
                .get(NS_STYLE, &format!("{prefix}-type"))
                .unwrap_or("single");
            let width = cascade
                .get(NS_STYLE, &format!("{prefix}-width"))
                .unwrap_or("auto");
            let color = cascade
                .get(NS_STYLE, &format!("{prefix}-color"))
                .unwrap_or("font-color");
            let mode = cascade
                .get(NS_STYLE, &format!("{prefix}-mode"))
                .unwrap_or("continuous");
            if style != "solid" || width != "auto" || color != "font-color" || mode != "continuous"
            {
                return Err(unsupported("text-decoration"));
            }
            match kind {
                "single" => Ok(Some("single")),
                "double" => Ok(Some("double")),
                _ => Err(unsupported("text-decoration")),
            }
        };
        let underline = line("text-underline")?;
        let strike = line("text-line-through")?;
        if cascade
            .get(NS_STYLE, "text-line-through-text")
            .is_some_and(|v| !v.is_empty())
            || cascade
                .get(NS_STYLE, "text-overline-style")
                .is_some_and(|v| v != "none")
        {
            return Err(unsupported("text-decoration"));
        }
        let background = match cascade.get(NS_FO, "background-color") {
            None | Some("transparent") => None,
            Some(value) => Some(hex_color(value)?),
        };
        // §20.395 automatic colour is chosen against the background. With no
        // background it is the dark default on a white page, as is DOCX
        // automatic colour; other backgrounds are not yet measured.
        let color_auto = cascade.get(NS_STYLE, "use-window-font-color") == Some("true");
        if color_auto && background.is_some() {
            return Err(unsupported("automatic-font-color"));
        }
        let color = if color_auto {
            None
        } else {
            cascade.get(NS_FO, "color").map(hex_color).transpose()?
        };
        let letter_spacing = match cascade.get(NS_FO, "letter-spacing") {
            None | Some("normal") => None,
            Some(value) => Some(length_pt(value)?),
        };
        let scale = cascade
            .get(NS_STYLE, "text-scale")
            .map(percent)
            .transpose()?
            .filter(|scale| (*scale - 1.0).abs() > 1e-9);
        let (position, size) = match cascade.get(NS_STYLE, "text-position") {
            None => (None, size),
            Some(value) => {
                let mut parts = value.split_whitespace();
                let raise = parts.next().ok_or_else(|| invalid("empty text-position"))?;
                // §20.384 leaves the keyword offsets and an omitted size to
                // the consumer: not yet measured.
                let relative = parts
                    .next()
                    .ok_or_else(|| unsupported("text-position-default"))?;
                if !raise.ends_with('%') {
                    return Err(unsupported("text-position-default"));
                }
                let raise = percent(raise)?;
                (
                    (raise != 0.0).then_some(raise * size),
                    size * percent(relative)?,
                )
            }
        };
        let caps = match cascade.get(NS_FO, "text-transform") {
            None | Some("none") => false,
            Some("uppercase") => true,
            Some(_) => return Err(unsupported("text-transform")),
        };
        if cascade
            .get(NS_FO, "font-variant")
            .is_some_and(|v| v != "normal")
        {
            // DOCX synthesizes small caps its own way; not yet measured.
            return Err(unsupported("small-caps"));
        }
        // §20.315 pair kerning on: DOCX kerning with a 0pt threshold kerns
        // every size (wordKerningApplies, outside compatibility mode 15).
        let kerning = match cascade.get(NS_STYLE, "letter-kerning") {
            Some("true") => Some(0.0),
            _ => return Err(unsupported("letter-kerning-off")),
        };
        let language = |language: &str, country: &str| -> Option<String> {
            let language = cascade.get(NS_STYLE, language).filter(|v| *v != "none")?;
            Some(
                match cascade.get(NS_STYLE, country).filter(|v| *v != "none") {
                    Some(country) => format!(
                        "{}-{}",
                        language.to_ascii_lowercase(),
                        country.to_ascii_uppercase()
                    ),
                    None => language.to_ascii_lowercase(),
                },
            )
        };
        Ok((
            Self {
                family,
                size,
                bold,
                italic,
                underline,
                strike,
                color,
                color_auto,
                background,
                letter_spacing,
                scale,
                position,
                caps,
                kerning,
                lang_east_asia: language("language-asian", "country-asian"),
                lang_bidi: language("language-complex", "country-complex"),
            },
            face,
        ))
    }

    fn text_run(&self) -> TextRun {
        let family = Some(self.family.clone());
        TextRun {
            bold: self.bold,
            italic: self.italic,
            bold_cs: Some(self.bold),
            italic_cs: Some(self.italic),
            underline: self.underline.is_some(),
            underline_style: self
                .underline
                .filter(|s| *s != "single")
                .map(str::to_string),
            strikethrough: self.strike == Some("single"),
            double_strikethrough: self.strike == Some("double"),
            font_size: self.size,
            font_size_cs: Some(self.size),
            color: self.color.clone(),
            color_auto: self.color_auto,
            // Every font slot carries the segment's face, so the DOCX slot
            // selection (ECMA-376 §17.3.2.26) cannot pick another face.
            font_family: family.clone(),
            font_family_high_ansi: family.clone(),
            font_family_east_asia: family.clone(),
            font_family_cs: family,
            background: self.background.clone(),
            char_spacing: self.letter_spacing,
            char_scale: self.scale,
            position: self.position,
            all_caps: self.caps,
            kerning: self.kerning,
            lang_east_asia: self.lang_east_asia.clone(),
            lang_bidi: self.lang_bidi.clone(),
            ..TextRun::default()
        }
    }

    /// Whether `run` was produced from identical properties.
    fn continues(&self, run: &TextRun) -> bool {
        let mut probe = self.text_run();
        probe.text.clone_from(&run.text);
        format!("{probe:?}") == format!("{run:?}")
    }

    fn font_facts(&self) -> RunFontFacts {
        let family = Some(self.family.clone());
        RunFontFacts {
            font_family: family.clone(),
            font_family_high_ansi: family.clone(),
            font_family_east_asia: family.clone(),
            font_family_cs: family,
            font_size: Some(self.size),
            font_size_cs: Some(self.size),
            bold: self.bold,
            italic: self.italic,
            bold_cs: Some(self.bold),
            italic_cs: Some(self.italic),
            kerning: self.kerning,
            ..RunFontFacts::default()
        }
    }
}

/// Text property keys the reader understands; any other key set anywhere in
/// a run's cascade must be inert, or the document is rejected.
fn check_text_keys(cascade: &Cascade<'_>) -> Result<(), String> {
    let hyphenation_off = cascade.get(NS_FO, "hyphenate") != Some("true");
    if !hyphenation_off {
        return Err(unsupported("hyphenation"));
    }
    for (ns, name) in cascade.keys() {
        let value = cascade.get(&ns, &name).unwrap_or("");
        let admitted = match (ns.as_str(), name.as_str()) {
            (NS_STYLE, n)
                if n.starts_with("font-name")
                    || n.starts_with("font-family")
                    || n.starts_with("font-pitch")
                    || n.starts_with("font-charset")
                    || n.starts_with("font-size")
                    || n.starts_with("font-weight")
                    || (n.starts_with("font-style") && !n.starts_with("font-style-name")) =>
            {
                true
            }
            (
                NS_FO,
                "font-family" | "font-size" | "font-weight" | "font-style" | "color"
                | "background-color" | "letter-spacing" | "text-transform" | "font-variant"
                | "language" | "country" | "script" | "hyphenate",
            ) => true,
            (
                NS_STYLE,
                "use-window-font-color"
                | "text-underline-style"
                | "text-underline-type"
                | "text-underline-width"
                | "text-underline-color"
                | "text-underline-mode"
                | "text-line-through-style"
                | "text-line-through-type"
                | "text-line-through-width"
                | "text-line-through-color"
                | "text-line-through-mode"
                | "text-line-through-text"
                | "text-overline-style"
                | "text-scale"
                | "text-position"
                | "letter-kerning"
                | "language-asian"
                | "country-asian"
                | "language-complex"
                | "country-complex"
                | "script-asian"
                | "script-complex",
            ) => true,
            // Hyphenation parameters are inert while fo:hyphenate is false.
            (NS_FO | NS_LOEXT, n) if n.starts_with("hyphenation-") => true,
            // Inert values of effects the reader does not render.
            (NS_STYLE, "font-relief") => value == "none",
            (NS_STYLE, "text-outline") => value == "false",
            (NS_FO, "text-shadow") => value == "none",
            (NS_STYLE, "text-emphasize") => value == "none",
            (NS_STYLE, "text-combine") => value == "none",
            (NS_STYLE, "text-rotation-angle") => value == "0" || value == "0deg",
            (NS_STYLE, "text-blinking") => value == "false",
            (NS_LOEXT, "opacity") => value == "0%",
            (NS_LOEXT, "font-optical-sizing") => value == "auto",
            (NS_OFFICEOOO, "rsid" | "paragraph-rsid") => true,
            _ => false,
        };
        if !admitted {
            return Err(unsupported("text-property"));
        }
    }
    Ok(())
}

/// The face of a script slot: the nearest style that sets either the face
/// reference (`style:font-name*`) or a direct family (`fo:font-family`,
/// `style:font-family-asian/-complex` with its generic/pitch attributes);
/// within one style the face reference wins.
fn resolve_family(cascade: &Cascade<'_>, suffix: &str, fonts: &FontFaces) -> Result<Face, String> {
    let name_key = (NS_STYLE.to_string(), format!("font-name{suffix}"));
    let family_key = if suffix.is_empty() {
        (NS_FO.to_string(), "font-family".to_string())
    } else {
        (NS_STYLE.to_string(), format!("font-family{suffix}"))
    };
    for style in &cascade.styles {
        let Some(group) = style.group(cascade.group) else {
            continue;
        };
        if let Some(name) = group.attributes.get(&name_key) {
            return fonts
                .faces
                .get(name)
                .cloned()
                .ok_or_else(|| invalid("a style names a missing font face"));
        }
        if let Some(family) = group.attributes.get(&family_key) {
            let attribute = |name: &str| {
                group
                    .attributes
                    .get(&(NS_STYLE.to_string(), format!("{name}{suffix}")))
            };
            return Ok(Face {
                family: unquote_family(family)?,
                generic: generic_class(attribute("font-family-generic").map(String::as_str))?,
                pitch: pitch(attribute("font-pitch").map(String::as_str))?,
                charset: charset(attribute("font-charset").map(String::as_str))?,
            });
        }
    }
    Err(unsupported("font-unresolved"))
}

/// Paragraph properties: exact mappings, inert values, rejection otherwise.
fn paragraph_properties(cascade: &Cascade<'_>) -> Result<Collected, String> {
    for (ns, name) in cascade.keys() {
        let value = cascade.get(&ns, &name).unwrap_or("");
        let admitted = match (ns.as_str(), name.as_str()) {
            (
                NS_FO,
                "margin-left" | "margin-right" | "margin-top" | "margin-bottom" | "text-indent"
                | "line-height" | "text-align" | "keep-with-next" | "keep-together" | "widows"
                | "orphans" | "break-before" | "break-after",
            ) => true,
            (
                NS_STYLE,
                "line-height-at-least" | "contextual-spacing" | "tab-stop-distance" | "page-number",
            ) => true,
            (NS_STYLE, "line-spacing") => length_pt(value).is_ok_and(|v| v == 0.0),
            (
                NS_STYLE,
                "auto-text-indent"
                | "justify-single-word"
                | "register-true"
                | "font-independent-line-spacing",
            ) => value == "false",
            (NS_FO, "text-align-last") => value == "start",
            (NS_STYLE, "writing-mode") => matches!(value, "lr-tb" | "page"),
            (NS_STYLE, "vertical-align") => matches!(value, "auto" | "baseline"),
            (NS_STYLE, "shadow") => value == "none",
            (NS_FO, "background-color") => value == "transparent",
            (NS_FO, "border" | "border-top" | "border-bottom" | "border-left" | "border-right") => {
                value == "none"
            }
            // Padding and border joining are inert without borders.
            (NS_FO, n) if n.starts_with("padding") => true,
            (NS_STYLE, "join-border") => true,
            (NS_FO | NS_LOEXT, n) if n.starts_with("hyphenation-") => true,
            // Inert without a layout grid (pages admit grid mode none only).
            (NS_STYLE, "snap-to-layout-grid") => true,
            // Line numbering is rejected document-wide when enabled.
            (NS_TEXT, "number-lines" | "line-number") => true,
            // Asian line breaking, punctuation wrapping and autospace affect
            // Asian-class text only, which is rejected.
            (NS_STYLE, "punctuation-wrap" | "line-break" | "text-autospace") => true,
            (NS_DRAW, "fill") => value == "none",
            (NS_OFFICEOOO, "rsid" | "paragraph-rsid") => true,
            _ => false,
        };
        if !admitted {
            return Err(unsupported("paragraph-property"));
        }
    }
    for element in cascade.element_names() {
        let entries = cascade.element(&element).unwrap_or(&[]);
        let empty = entries
            .iter()
            .all(|e| e.children.is_empty() && e.attributes.is_empty());
        match element.as_str() {
            "tab-stops" if entries.iter().all(|e| e.children.is_empty()) => {}
            "background-image" if empty => {}
            "tab-stops" => return Err(unsupported("tab-stops")),
            _ => return Err(unsupported("paragraph-property")),
        }
    }
    let length = |name: &str| -> Result<f64, String> {
        match cascade.get(NS_FO, name) {
            None => Ok(0.0),
            Some(value) if value.ends_with('%') => Err(unsupported("relative-margin")),
            Some(value) => length_pt(value),
        }
    };
    let alignment = match cascade.get(NS_FO, "text-align") {
        None | Some("start") | Some("left") => "left",
        Some("end") | Some("right") => "right",
        Some("center") => "center",
        Some("justify") => "justify",
        Some(_) => return Err(invalid("bad text-align")),
    };
    // One active line-height control (§20.204, §20.317, §20.318).
    let line_spacing = match (
        cascade.get(NS_FO, "line-height"),
        cascade.get(NS_STYLE, "line-height-at-least"),
    ) {
        (Some("normal"), _) | (None, None) => None,
        (Some(value), None) if value.ends_with('%') => Some(LineSpacing {
            value: percent(value)? * 240.0,
            rule: "auto".into(),
            explicit: true,
        }),
        (Some(value), None) => Some(LineSpacing {
            value: length_pt(value)?,
            rule: "exact".into(),
            explicit: true,
        }),
        (None, Some(value)) => Some(LineSpacing {
            value: length_pt(value)?,
            rule: "atLeast".into(),
            explicit: true,
        }),
        (Some(_), Some(_)) => return Err(unsupported("line-height-combination")),
    };
    let count = |name: &str| -> Result<u32, String> {
        Ok(cascade
            .get(NS_FO, name)
            .map(|v| v.parse::<u32>().map_err(|_| invalid("bad widows/orphans")))
            .transpose()?
            .unwrap_or(2))
    };
    let widow_control = match (count("widows")?, count("orphans")?) {
        (2, 2) => true,
        (a, b) if a <= 1 && b <= 1 => false,
        _ => return Err(unsupported("widows-orphans")),
    };
    let breaking = |name: &str| -> Result<bool, String> {
        match cascade.get(NS_FO, name) {
            None | Some("auto") => Ok(false),
            Some("page") => Ok(true),
            Some(_) => Err(unsupported("column-break")),
        }
    };
    let paragraph = DocParagraph {
        alignment: alignment.into(),
        indent_left: length("margin-left")?,
        indent_right: length("margin-right")?,
        indent_first: length("text-indent")?,
        line_spacing,
        keep_next: cascade.get(NS_FO, "keep-with-next") == Some("always"),
        keep_lines: cascade.get(NS_FO, "keep-together") == Some("always"),
        widow_control,
        page_break_before: breaking("break-before")?,
        ..DocParagraph::default()
    };
    Ok(Collected {
        paragraph,
        top: length("margin-top")?,
        bottom: length("margin-bottom")?,
        contextual: cascade.get(NS_STYLE, "contextual-spacing") == Some("true"),
        named_style: None,
        page_break_after: breaking("break-after")?,
    })
}

/// ODF lengths (Part 3 §18.3.18): cm, mm, in, pt, pc, px (1/96 in).
pub fn length_pt(value: &str) -> Result<f64, String> {
    let value = value.trim();
    let split = value
        .find(|c: char| c.is_ascii_alphabetic())
        .ok_or_else(|| invalid("a length has no unit"))?;
    let (number, unit) = value.split_at(split);
    let number: f64 = number
        .parse()
        .map_err(|_| invalid("a length is not a number"))?;
    let factor = match unit {
        "pt" => 1.0,
        "cm" => 72.0 / 2.54,
        "mm" => 72.0 / 25.4,
        "in" => 72.0,
        "pc" => 12.0,
        "px" => 0.75,
        _ => return Err(invalid("a length has an unknown unit")),
    };
    let points = number * factor;
    if !points.is_finite() || points.abs() > 1.0e6 {
        return Err(invalid("a length is out of range"));
    }
    Ok(points)
}

fn percent(value: &str) -> Result<f64, String> {
    let number = value
        .trim()
        .strip_suffix('%')
        .ok_or_else(|| invalid("a percentage has no %"))?;
    let number: f64 = number
        .parse()
        .map_err(|_| invalid("a percentage is not a number"))?;
    if !number.is_finite() || !(-1000.0..=10000.0).contains(&number) {
        return Err(invalid("a percentage is out of range"));
    }
    Ok(number / 100.0)
}

fn hex_color(value: &str) -> Result<String, String> {
    let hex = value
        .strip_prefix('#')
        .filter(|h| h.len() == 6 && h.chars().all(|c| c.is_ascii_hexdigit()))
        .ok_or_else(|| invalid("a colour is not #rrggbb"))?;
    Ok(hex.to_ascii_lowercase())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::tests::{manifest, package, ODT};
    use crate::package::PackageLimits;
    use ooxml_common::resource::OoxmlFormat;

    const NS: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0""#;
    const FACES: &str = r#"<office:font-face-decls><style:font-face style:name="Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman"/><style:font-face style:name="Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss"/><style:font-face style:name="Sans2" svg:font-family="'Liberation Sans'" style:font-family-generic="system"/><style:font-face style:name="Mincho" svg:font-family="'Hiragino Mincho ProN'" style:font-family-generic="system"/></office:font-face-decls>"#;

    struct Fixture<'a> {
        default_text: &'a str,
        styles: &'a str,
        automatic: &'a str,
        master: &'a str,
        body: &'a str,
    }

    impl Default for Fixture<'_> {
        fn default() -> Self {
            Self {
                default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Serif" style:font-size-asian="12pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true""#,
                styles: "",
                automatic: "",
                master: "",
                body: "<text:p>x</text:p>",
            }
        }
    }

    fn read_fixture(fixture: Fixture<'_>) -> Result<Document, String> {
        let styles = format!(
            r#"<office:document-styles {NS} office:version="1.3">{FACES}<office:styles><style:default-style style:family="paragraph"><style:text-properties {}/></style:default-style><style:default-style style:family="graphic"/>{}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="210mm" fo:page-height="297mm" fo:margin-top="20mm" fo:margin-bottom="20mm" fo:margin-left="25mm" fo:margin-right="25mm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">{}</style:master-page></office:master-styles></office:document-styles>"#,
            fixture.default_text, fixture.styles, fixture.master
        );
        let content = format!(
            r#"<office:document-content {NS} office:version="1.3">{FACES}<office:automatic-styles>{}</office:automatic-styles><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            fixture.automatic, fixture.body
        );
        let bytes = package(ODT, &[
            ("META-INF/manifest.xml", &manifest(ODT, "")),
            ("styles.xml", &styles),
            ("content.xml", &content),
        ]);
        let package = OdfPackage::open(bytes, OoxmlFormat::Docx, PackageLimits::default(), &MEDIA_TYPES)?;
        read(&package)
    }

    fn paragraphs(document: &Document) -> Vec<&DocParagraph> {
        document
            .body
            .iter()
            .map(|element| match element {
                BodyElement::Paragraph(paragraph) => paragraph.as_ref(),
                _ => panic!("not a paragraph"),
            })
            .collect()
    }

    fn texts(paragraph: &DocParagraph) -> Vec<String> {
        paragraph
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(text) => text.text.clone(),
                DocRun::Break { .. } => "<br>".into(),
                _ => "?".into(),
            })
            .collect()
    }

    #[test]
    fn collapses_white_space_and_keeps_explicit_spaces_tabs_and_breaks() {
        let document = read_fixture(Fixture {
            body: "<text:p>\n  a \t b<text:s text:c=\"2\"/>c<text:tab/>d<text:line-break/>  e  </text:p>",
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(texts(paragraphs(&document)[0]), ["a b  c\td", "<br>", " e"]);
        assert!((document.section.margin_left - 70.866).abs() < 0.01);
    }

    #[test]
    fn adds_adjacent_margins_and_removes_them_between_contextual_paragraphs() {
        let styles = r#"<style:style style:name="A" style:family="paragraph"><style:paragraph-properties fo:margin-top="6pt" fo:margin-bottom="10pt"/></style:style><style:style style:name="C" style:family="paragraph"><style:paragraph-properties fo:margin-top="12pt" fo:margin-bottom="12pt" style:contextual-spacing="true"/></style:style>"#;
        let document = read_fixture(Fixture {
            styles,
            body: r#"<text:p text:style-name="A">1</text:p><text:p text:style-name="A">2</text:p><text:p text:style-name="C">3</text:p><text:p text:style-name="C">4</text:p><text:p text:style-name="A">5</text:p>"#,
            ..Fixture::default()
        })
        .unwrap();
        let gaps: Vec<f64> = paragraphs(&document).iter().map(|p| p.space_before).collect();
        assert_eq!(gaps, [6.0, 16.0, 22.0, 0.0, 18.0]);
        assert!(paragraphs(&document).iter().all(|p| p.space_after == 0.0 && !p.contextual_spacing));

        let mixed = read_fixture(Fixture {
            styles,
            automatic: r#"<style:style style:name="P1" style:family="paragraph" style:parent-style-name="C"><style:paragraph-properties fo:margin-top="12pt" fo:margin-bottom="12pt" style:contextual-spacing="false"/></style:style>"#,
            body: r#"<text:p text:style-name="C">1</text:p><text:p text:style-name="P1">2</text:p>"#,
            ..Fixture::default()
        });
        assert_eq!(mixed.unwrap_err(), unsupported("contextual-spacing-mixed"));
    }

    #[test]
    fn maps_line_height_controls_and_rejects_their_combination() {
        let style = |props: &str| format!(r#"<style:style style:name="L" style:family="paragraph"><style:paragraph-properties {props}/></style:style>"#);
        let spacing = |props: &str| {
            read_fixture(Fixture { styles: &style(props), body: r#"<text:p text:style-name="L">x</text:p>"#, ..Fixture::default() })
                .map(|d| paragraphs(&d)[0].line_spacing.as_ref().map(|s| (s.value, s.rule.clone())))
        };
        assert_eq!(spacing(r#"fo:line-height="115%""#).unwrap(), Some((276.0, "auto".into())));
        assert_eq!(spacing(r#"fo:line-height="20pt""#).unwrap(), Some((20.0, "exact".into())));
        assert_eq!(spacing(r#"style:line-height-at-least="30pt""#).unwrap(), Some((30.0, "atLeast".into())));
        assert_eq!(spacing(r#"fo:line-height="normal" style:line-height-at-least="30pt""#).unwrap(), None);
        assert_eq!(
            spacing(r#"fo:line-height="20pt" style:line-height-at-least="30pt""#).unwrap_err(),
            unsupported("line-height-combination")
        );
        assert_eq!(spacing(r#"style:line-spacing="2pt""#).unwrap_err(), unsupported("paragraph-property"));
    }

    #[test]
    fn resolves_span_properties_through_the_cascade() {
        let document = read_fixture(Fixture {
            styles: r#"<style:style style:name="Big" style:family="paragraph"><style:text-properties fo:font-size="20pt"/></style:style>"#,
            automatic: r#"<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" style:font-name="Sans"/></style:style><style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"#,
            body: r#"<text:p text:style-name="Big">a<text:span text:style-name="T1">b<text:span text:style-name="T2">c</text:span></text:span></text:p>"#,
            ..Fixture::default()
        })
        .unwrap();
        let runs: Vec<&TextRun> = paragraphs(&document)[0]
            .runs
            .iter()
            .map(|r| match r { DocRun::Text(t) => t.as_ref(), _ => panic!() })
            .collect();
        assert_eq!((runs[0].font_size, runs[0].bold, runs[0].font_family.as_deref()), (20.0, false, Some("Liberation Serif")));
        assert_eq!((runs[1].bold, runs[1].italic, runs[1].font_family.as_deref()), (true, false, Some("Liberation Sans")));
        assert_eq!((runs[2].bold, runs[2].italic, runs[2].font_size), (true, true, 20.0));
        // Only used faces enter the font table; the unused "Sans2" twin with a
        // different generic class does not conflict.
        assert_eq!(document.font_family_classes.get("Liberation Sans").map(String::as_str), Some("swiss"));
    }

    #[test]
    fn rejects_unadmitted_content_with_reason_codes() {
        let reject = |fixture: Fixture<'_>| read_fixture(fixture).unwrap_err();
        assert_eq!(reject(Fixture { body: "<text:list/>", ..Fixture::default() }), unsupported("body-element"));
        assert_eq!(reject(Fixture { body: "<text:p>日本</text:p>", ..Fixture::default() }), unsupported("script-class"));
        assert_eq!(reject(Fixture { master: "<style:header><text:p>h</text:p></style:header>", ..Fixture::default() }), unsupported("header-footer"));
        assert_eq!(
            reject(Fixture { default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:letter-kerning="false""#, ..Fixture::default() }),
            unsupported("letter-kerning-off")
        );
        assert_eq!(
            reject(Fixture {
                automatic: r#"<style:style style:name="T" style:family="text"><style:text-properties style:text-outline="true"/></style:style>"#,
                body: r#"<text:p><text:span text:style-name="T">x</text:span></text:p>"#,
                ..Fixture::default()
            }),
            unsupported("text-property")
        );
        // A paragraph-initial character outside Table 22 whose script classes
        // resolve to different faces.
        assert_eq!(
            reject(Fixture {
                default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="12pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true""#,
                body: "<text:p>\u{2014}x</text:p>",
                ..Fixture::default()
            }),
            unsupported("script-initial-weak-character")
        );
        // Two used faces of one family with different metadata.
        assert_eq!(
            reject(Fixture {
                automatic: r#"<style:style style:name="A" style:family="text"><style:text-properties style:font-name="Sans"/></style:style><style:style style:name="B" style:family="text"><style:text-properties style:font-name="Sans2"/></style:style>"#,
                body: r#"<text:p><text:span text:style-name="A">a</text:span><text:span text:style-name="B">b</text:span></text:p>"#,
                ..Fixture::default()
            }),
            unsupported("font-face-conflict")
        );
    }

    #[test]
    fn weak_characters_follow_the_preceding_class() {
        let document = read_fixture(Fixture {
            default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="30pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true""#,
            body: "<text:p>a\u{2014}b c</text:p>",
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(texts(paragraphs(&document)[0]), ["a\u{2014}b c"]);
    }
}
