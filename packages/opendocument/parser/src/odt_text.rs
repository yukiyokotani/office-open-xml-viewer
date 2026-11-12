//! Inline content (ODF Part 3 §6), white-space processing (§6.1.2), script
//! classes (§20.358) and character properties (§20.180-§20.396).
//!
//! A paragraph's content is flattened into tokens, each referring to a
//! formatting context (a span's style chain, hyperlink and revision), then
//! collapsed and emitted as DOCX runs. Every ODF script class keeps its own
//! face, size, weight and posture: a run holds characters of one class and
//! carries that class's face in every DOCX font slot, so the DOCX slot
//! selection (ECMA-376 §17.3.2.26) cannot pick another face.

use super::*;
use crate::script::{is_weak_unmapped, table_22};
use crate::xml::NS_XML;
use docx_model::{EmbeddedFont, FieldRun, NoteRef, RubyAnnotation, RunBorder};

/// A DOCX kerning threshold (pt) below every font size: kerning applies at
/// any size (§17.3.2.19 `w:kern`; zero would disable it).
const KERN_EVERY_SIZE: f64 = f64::MIN_POSITIVE;

// ---------------------------------------------------------------------------
// settings.xml

/// The LibreOffice configuration items (`settings.xml`, not standardised by
/// ODF) the reader applies.
#[derive(Clone, Copy, Debug, Default)]
pub(super) struct LayoutFlags {
    /// DoNotJustifyLinesWithManualBreak: a justified line ended by a manual
    /// break is left unjustified (LibreOffice otherwise stretches it).
    pub no_justify_breaks: bool,
    /// BalanceSpacesAndIdeographicSpaces: a U+0020 at a text portion's edge
    /// or next to another U+0020 advances half its Asian font size.
    pub balance: bool,
    /// TabsRelativeToIndent: tab positions count from the paragraph indent
    /// when true (LibreOffice's default), from the text area when false.
    pub tabs_relative: Option<bool>,
    /// JustifyLinesWithShrinking: justification may also compress spaces
    /// (LibreOffice's emulation of Word's mode-15 justification).
    pub shrink: bool,
}

pub(super) fn read_settings(root: roxmltree::Node<'_, '_>) -> Result<LayoutFlags, String> {
    let mut flags = LayoutFlags::default();
    // LibreOffice reads its layout items from the top-level
    // "ooo:configuration-settings" set only.
    let Some(set) = root
        .kids()
        .filter(|n| xml::is(*n, NS_OFFICE, "settings"))
        .flat_map(|settings| settings.kids())
        .find(|n| {
            xml::is(*n, xml::NS_CONFIG, "config-item-set")
                && n.attribute((xml::NS_CONFIG, "name")) == Some("ooo:configuration-settings")
        })
    else {
        return Ok(flags);
    };
    for item in set
        .kids()
        .filter(|n| xml::is(*n, xml::NS_CONFIG, "config-item"))
    {
        budget::tick(1)?;
        let value = item.text().unwrap_or("").trim() == "true";
        match item.attribute((xml::NS_CONFIG, "name")) {
            Some("DoNotJustifyLinesWithManualBreak") => flags.no_justify_breaks = value,
            Some("BalanceSpacesAndIdeographicSpaces") => flags.balance = value,
            Some("TabsRelativeToIndent") => flags.tabs_relative = Some(value),
            Some("JustifyLinesWithShrinking") => flags.shrink = value,
            _ => {}
        }
    }
    Ok(flags)
}

// ---------------------------------------------------------------------------
// Font faces

/// `style:font-face` declarations (Part 3 §16.23): the face a style
/// references, its family and the metadata the DOCX font table carries.
#[derive(Default)]
pub(super) struct FontFaces {
    faces: HashMap<String, Face>,
    /// Interned property strings shared by every resolution that uses them.
    atoms: RefCell<HashSet<Rc<str>>>,
    /// Embedded font files (§16.23 `svg:font-face-src`), each package path
    /// once; the family is the face's shared atom.
    pub embedded: Vec<(Atom, String, &'static str)>,
    embedded_paths: HashSet<String>,
}

/// An interned string: equal atoms are the same allocation, so cached run
/// properties share them and compare and hash by identity in O(1).
#[derive(Clone, Debug)]
pub(super) struct Atom(Rc<str>);

impl PartialEq for Atom {
    fn eq(&self, other: &Self) -> bool {
        Rc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for Atom {}

impl std::hash::Hash for Atom {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        std::ptr::hash(Rc::as_ptr(&self.0) as *const u8, state);
    }
}

impl std::ops::Deref for Atom {
    type Target = str;
    fn deref(&self) -> &str {
        &self.0
    }
}

#[derive(Clone, PartialEq, Eq, Hash, Debug)]
pub(super) struct Face {
    pub family: Atom,
    pub generic: Option<&'static str>,
    pub pitch: Option<&'static str>,
    pub charset: Option<&'static str>,
}

impl FontFaces {
    pub fn atom(&self, value: &str) -> Atom {
        let mut atoms = self.atoms.borrow_mut();
        if let Some(existing) = atoms.get(value) {
            return Atom(existing.clone());
        }
        let atom: Rc<str> = value.into();
        atoms.insert(atom.clone());
        Atom(atom)
    }

    pub fn add(
        &mut self,
        decls: roxmltree::Node<'_, '_>,
        package: &OdfPackage,
    ) -> Result<(), String> {
        for face in decls.kids().filter(|n| xml::is(*n, NS_STYLE, "font-face")) {
            budget::tick(1)?;
            let Some(name) = face.attribute((NS_STYLE, "name")) else {
                continue;
            };
            let family = face
                .attribute((NS_SVG, "font-family"))
                .map(first_family)
                .unwrap_or_else(|| name.to_string());
            let value = Face {
                family: self.atom(&family),
                generic: generic_class(face.attribute((NS_STYLE, "font-family-generic"))),
                pitch: pitch(face.attribute((NS_STYLE, "font-pitch"))),
                charset: charset(face.attribute((NS_STYLE, "font-charset"))),
            };
            // Embedded font files (§16.23, svg:font-face-uri in the package).
            for uri in face
                .subtree()
                .filter(|n| xml::is(*n, NS_SVG, "font-face-uri"))
            {
                budget::tick(1)?;
                let Some(href) = uri.attribute((NS_XLINK, "href")) else {
                    continue;
                };
                if let Some(iri::Target::Package {
                    path,
                    fragment: None,
                }) = iri::resolve("content.xml", href)
                {
                    if package.contains(&path) && !self.embedded_paths.contains(&path) {
                        let style = match (
                            face.attribute((NS_SVG, "font-weight")).unwrap_or("normal"),
                            face.attribute((NS_SVG, "font-style")).unwrap_or("normal"),
                        ) {
                            ("bold" | "700" | "800" | "900", "italic" | "oblique") => "boldItalic",
                            ("bold" | "700" | "800" | "900", _) => "bold",
                            (_, "italic" | "oblique") => "italic",
                            _ => "regular",
                        };
                        self.embedded_paths.insert(path.clone());
                        self.embedded.push((value.family.clone(), path, style));
                    }
                }
            }
            self.faces.entry(name.to_string()).or_insert(value);
        }
        Ok(())
    }
}

/// The first family of a CSS font-family list, unquoted.
fn first_family(value: &str) -> String {
    let first = value.split(',').next().unwrap_or("").trim();
    first
        .strip_prefix('\'')
        .and_then(|v| v.strip_suffix('\''))
        .or_else(|| first.strip_prefix('"').and_then(|v| v.strip_suffix('"')))
        .unwrap_or(first)
        .to_string()
}

/// The DOCX font table (classes, pitches, charsets) is keyed by family name;
/// the first face of a family wins.
pub(super) fn used_font_table(used: &HashSet<Face>) -> HashMap<String, Face> {
    let mut faces: Vec<&Face> = used.iter().collect();
    faces.sort_by(|a, b| {
        (&*a.family, a.generic, a.pitch, a.charset)
            .cmp(&(&*b.family, b.generic, b.pitch, b.charset))
    });
    let mut table: HashMap<String, Face> = HashMap::new();
    for face in faces {
        table
            .entry(face.family.to_string())
            .or_insert_with(|| face.clone());
    }
    table
}

pub(super) fn metadata(
    table: &HashMap<String, Face>,
    field: impl Fn(&Face) -> Option<&'static str>,
) -> BTreeMap<String, String> {
    table
        .values()
        .filter_map(|face| field(face).map(|value| (face.family.to_string(), value.to_string())))
        .collect()
}

/// ODF `style:font-family-generic` (§20.278) and DOCX `w:family` (ECMA-376
/// §17.8.3.10) share their vocabulary; `system` has no DOCX class.
fn generic_class(value: Option<&str>) -> Option<&'static str> {
    match value {
        Some("roman") => Some("roman"),
        Some("swiss") => Some("swiss"),
        Some("modern") => Some("modern"),
        Some("decorative") => Some("decorative"),
        Some("script") => Some("script"),
        _ => None,
    }
}

fn pitch(value: Option<&str>) -> Option<&'static str> {
    match value {
        Some("fixed") => Some("fixed"),
        Some("variable") => Some("variable"),
        _ => None,
    }
}

/// §20.268 `style:font-charset` is an encoding name; `x-symbol` is the
/// symbol charset (02).
fn charset(value: Option<&str>) -> Option<&'static str> {
    match value {
        Some("x-symbol") => Some("02"),
        _ => None,
    }
}

impl Reader<'_> {
    /// Embedded font files as DOCX embedded fonts, retained under their
    /// package paths (DOCX reads a non-`.odttf` part as a plain font).
    pub(super) fn embedded_fonts(&mut self) -> Result<Vec<EmbeddedFont>, String> {
        let mut fonts = Vec::new();
        // Each package path is retained once.
        let faces = self.fonts;
        let mut retained: HashSet<&str> = HashSet::new();
        let mut added = Vec::new();
        for (family, path, style) in &faces.embedded {
            budget::tick(1)?;
            if !self.image_types.contains_key(path) && retained.insert(path.as_str()) {
                let bytes = self
                    .package
                    .session()
                    .run_operation("odf-font", |operation| operation.read_bytes(path))?;
                added.push(OwnedDocumentResource {
                    key: path.clone(),
                    mime_type: "font/ttf".to_string(),
                    bytes,
                });
            }
            fonts.push(EmbeddedFont {
                font_name: family.to_string(),
                style: style.to_string(),
                part_path: path.clone(),
                font_key: String::new(),
            });
        }
        drop(retained);
        self.images.extend(added);
        Ok(fonts)
    }
}

// ---------------------------------------------------------------------------
// Tracked changes (§5.5)

/// The change regions of `text:tracked-changes`: insertions are shown as
/// ordinary text and deletions are left out, the accepted state the DOCX
/// renderer projects for its own revisions.
pub(super) struct TrackedChanges<'a, 'i> {
    /// `text:id` → whether the region is a deletion.
    deletions: HashSet<&'a str>,
    _marker: std::marker::PhantomData<&'i ()>,
}

impl<'a, 'i> TrackedChanges<'a, 'i> {
    pub fn read(body: roxmltree::Node<'a, 'i>) -> Result<Self, String> {
        let mut deletions = HashSet::new();
        for log in body
            .kids()
            .filter(|n| xml::is(*n, NS_TEXT, "tracked-changes"))
        {
            for region in log
                .kids()
                .filter(|n| xml::is(*n, NS_TEXT, "changed-region"))
            {
                budget::tick(1)?;
                let id = region
                    .attribute((NS_TEXT, "id"))
                    .or_else(|| region.attribute((NS_XML, "id")));
                if let Some(id) = id {
                    if region.kids().any(|n| xml::is(n, NS_TEXT, "deletion")) {
                        deletions.insert(id);
                    }
                }
            }
        }
        Ok(Self {
            deletions,
            _marker: std::marker::PhantomData,
        })
    }

    fn is_deletion(&self, id: &str) -> bool {
        self.deletions.contains(id)
    }
}

// ---------------------------------------------------------------------------
// Inline content

/// One formatting context: a span's own style chain followed, through
/// `parent`, by the enclosing contexts (§16.2), plus the hyperlink that
/// encloses it and whether its text is hidden.
pub(super) struct TextContext<'a> {
    pub styles: Rc<[StyleRef<'a>]>,
    pub parent: Option<u32>,
    pub link: Option<Rc<(Option<String>, Option<String>)>>,
}

/// One paragraph's flattened inline content.
pub(super) struct Inline<'a> {
    pub contexts: Vec<TextContext<'a>>,
    by_span: HashMap<(u32, String), u32>,
    pub tokens: Vec<Token>,
    pub items: Vec<Item>,
    /// Characters of a span inside a deleted change region are dropped.
    in_deletion: bool,
    /// Inside a fieldmark's command (before its separator).
    in_field_code: u32,
    /// Bookmarks named in the paragraph.
    pub bookmarks: Vec<String>,
}

impl<'a> Inline<'a> {
    pub fn new(chain: Rc<[StyleRef<'a>]>) -> Self {
        Self {
            contexts: vec![TextContext {
                styles: chain,
                parent: None,
                link: None,
            }],
            by_span: HashMap::new(),
            tokens: Vec::new(),
            items: Vec::new(),
            in_deletion: false,
            in_field_code: 0,
            bookmarks: Vec::new(),
        }
    }
}

#[derive(Clone, Copy)]
pub(super) enum Token {
    Char(char, u32),
    /// `text:s`: one space that is never collapsed.
    Space(u32),
    Tab(u32),
    LineBreak,
    /// No character, but where LibreOffice may end its text portion
    /// (BalanceSpacesAndIdeographicSpaces): `true` for a bookmark or
    /// reference mark, `false` for an empty formatted span.
    Boundary(bool),
    /// An object or field (index into `Inline::items`) in a context.
    Item(u32, u32),
}

pub(super) enum Item {
    /// A prebuilt run (picture, shape).
    Run(DocRun),
    /// A field whose value the DOCX layout computes (page number, count).
    Field {
        kind: &'static str,
        instruction: &'static str,
        fallback: String,
    },
    /// A note citation in the text.
    NoteReference {
        kind: &'static str,
        id: String,
        label: Option<String>,
    },
    Ruby {
        base: String,
        text: String,
        /// The ruby text's size as its style sets it (`size_terms`); a
        /// missing base is the base text's size.
        ruby_size: Option<(Option<f64>, f64, f64)>,
    },
}

impl<'a> Reader<'a> {
    /// Inline content (§6.1): character data, spans, white-space elements,
    /// links, fields, notes, rubies and anchored drawings, flattened into
    /// tokens for §6.1.2 collapsing.
    pub(super) fn inline(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
        depth: usize,
    ) -> Result<(), String> {
        if depth > MAX_INLINE_DEPTH {
            return Ok(());
        }
        for child in node.kids() {
            budget::tick(1)?;
            if child.is_text() {
                if inline.in_deletion || inline.in_field_code > 0 {
                    continue;
                }
                let text = child.text().unwrap_or("");
                inline
                    .tokens
                    .extend(text.chars().map(|c| Token::Char(c, context)));
                self.admit_chars(0, inline)?;
                continue;
            }
            if !child.is_element() {
                continue;
            }
            // Field-code text of a fieldmark (its command) is not shown.
            let (namespace, name) = (child.tag_name().namespace(), child.tag_name().name());
            if namespace == Some(xml::NS_FIELD) {
                match name {
                    "fieldmark-start" => inline.in_field_code += 1,
                    "fieldmark-separator" | "fieldmark-end" => {
                        inline.in_field_code = inline.in_field_code.saturating_sub(1)
                    }
                    "fieldmark" => self.checkbox_fieldmark(child, context, inline)?,
                    _ => {}
                }
                continue;
            }
            if inline.in_field_code > 0 {
                continue;
            }
            // Tracked changes (§5.5): `text:change-start`/`-end` bracket a
            // region; deletions leave the accepted text without it.
            if namespace == Some(NS_TEXT)
                && matches!(name, "change-start" | "change-end" | "change")
            {
                let id = child.attribute((NS_TEXT, "change-id")).unwrap_or("");
                if self.changes.is_deletion(id) && name == "change-start" {
                    inline.in_deletion = true;
                } else if name == "change-end" && self.changes.is_deletion(id) {
                    inline.in_deletion = false;
                }
                continue;
            }
            if inline.in_deletion {
                continue;
            }
            match (namespace, name) {
                (Some(NS_TEXT), "span") => {
                    let nested = self.span_context(child, context, inline)?;
                    let before = inline.tokens.len();
                    self.inline(child, nested, inline, depth + 1)?;
                    if self.settings.balance && inline.tokens.len() == before && nested != context {
                        self.admit_chars(1, inline)?;
                        inline.tokens.push(Token::Boundary(false));
                    }
                }
                // §6.1.3 text:s.
                (Some(NS_TEXT), "s") => {
                    // `text:c` (§19.763) defaults to 1; a count over the
                    // paragraph bound rejects. The run is written in
                    // pieces, a step each, never allocated at once.
                    let count = xml::bounded_count(
                        child.attribute((NS_TEXT, "c")),
                        MAX_PARAGRAPH_CHARS as u64,
                        "paragraph character",
                    )?
                    .map_or(1, |count| count as usize);
                    self.admit_chars(count, inline)?;
                    let mut left = count;
                    while left > 0 {
                        let piece = left.min(4096);
                        budget::tick(1)?;
                        inline
                            .tokens
                            .extend(std::iter::repeat_n(Token::Space(context), piece));
                        left -= piece;
                    }
                }
                (Some(NS_TEXT), "tab") => {
                    self.admit_chars(1, inline)?;
                    inline.tokens.push(Token::Tab(context));
                }
                (Some(NS_TEXT), "line-break") => {
                    self.admit_chars(1, inline)?;
                    inline.tokens.push(Token::LineBreak);
                }
                (Some(NS_TEXT), "soft-page-break") => {}
                (Some(NS_TEXT), "bookmark" | "bookmark-start") => {
                    if let Some(name) = child.attribute((NS_TEXT, "name")) {
                        inline.bookmarks.push(name.to_string());
                    }
                    self.admit_chars(1, inline)?;
                    inline.tokens.push(Token::Boundary(true));
                }
                (
                    Some(NS_TEXT),
                    "bookmark-end"
                    | "reference-mark"
                    | "reference-mark-start"
                    | "reference-mark-end",
                ) => {
                    self.admit_chars(1, inline)?;
                    inline.tokens.push(Token::Boundary(true));
                }
                // §6.6 hyperlinks: the link's own text style (if any)
                // formats its content.
                (Some(NS_TEXT), "a") => {
                    let nested = self.link_context(child, context, inline)?;
                    self.inline(child, nested, inline, depth + 1)?;
                }
                (Some(NS_TEXT), "note") => self.note(child, context, inline)?,
                (Some(NS_TEXT), "ruby") => self.ruby(child, context, inline)?,
                // §7.3 page number and §7.5.17 page count fields.
                (Some(NS_TEXT), "page-number") => {
                    if child.attribute((NS_TEXT, "select-page")) == Some("previous")
                        || child.attribute((NS_TEXT, "select-page")) == Some("next")
                    {
                        // An adjacent page's number: its displayed text.
                        self.inline(child, context, inline, depth + 1)?;
                    } else {
                        self.field(child, "page", "PAGE", context, inline)?;
                    }
                }
                (Some(NS_TEXT), "page-count") => {
                    self.field(child, "numPages", "NUMPAGES", context, inline)?
                }
                // Fields whose content is a hidden value (§7.4, §7.7).
                (
                    Some(NS_TEXT),
                    "variable-set" | "user-field-input" | "sequence-ref" | "hidden-text",
                ) if child.attribute((NS_TEXT, "display")) == Some("none")
                    || child.attribute((NS_TEXT, "is-hidden")) == Some("true") => {}
                (Some(NS_TEXT), "hidden-paragraph") => {}
                (Some(NS_TEXT), _) if child.attribute((NS_TEXT, "display")) == Some("none") => {}
                // Index marks, change marks and the annotation (comment) are
                // not displayed text.
                (
                    Some(NS_TEXT),
                    "toc-mark"
                    | "toc-mark-start"
                    | "toc-mark-end"
                    | "alphabetical-index-mark"
                    | "alphabetical-index-mark-start"
                    | "alphabetical-index-mark-end"
                    | "user-index-mark"
                    | "user-index-mark-start"
                    | "user-index-mark-end"
                    | "bibliography-mark",
                ) => {
                    if name == "bibliography-mark" {
                        self.inline(child, context, inline, depth + 1)?;
                    }
                }
                (Some(NS_OFFICE), "annotation" | "annotation-end") => {}
                // Every other text field (§7) shows its current value, which
                // the producer stores as the element's content; meta
                // elements and content controls hold ordinary text.
                (Some(NS_TEXT), _) | (Some(NS_LOEXT), _) => {
                    self.inline(child, context, inline, depth + 1)?;
                }
                // Drawing shapes and frames anchored in the paragraph (§9).
                (Some(NS_DRAW), _) => {
                    if let Some(run) = self.drawing(child, Some((context, &*inline)))? {
                        let index = inline.items.len() as u32;
                        inline.items.push(Item::Run(run));
                        self.admit_chars(1, inline)?;
                        inline.tokens.push(Token::Item(index, context));
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    /// A checkbox fieldmark (LibreOffice's form field): its box state.
    fn checkbox_fieldmark(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<(), String> {
        let mut checked = false;
        for param in node.kids() {
            budget::tick(1)?;
            checked |= xml::is(param, xml::NS_FIELD, "param")
                && param.attribute((xml::NS_FIELD, "name"))
                    == Some("vnd.oasis.opendocument.field.checked")
                && param.attribute((xml::NS_FIELD, "value")) == Some("true");
        }
        self.admit_chars(1, inline)?;
        inline
            .tokens
            .push(Token::Char(if checked { '☒' } else { '☐' }, context));
        Ok(())
    }

    fn span_context(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<u32, String> {
        let Some(name) = node.attribute((NS_TEXT, "style-name")) else {
            return Ok(context);
        };
        if let Some(index) = inline.by_span.get(&(context, name.to_string())) {
            return Ok(*index);
        }
        let Some(style) = self.find_style(Family::Text, name) else {
            return Ok(context);
        };
        if inline.contexts.len() >= MAX_PARAGRAPH_CONTEXTS {
            return Ok(context);
        }
        let styles = self.style_chain(style)?;
        let link = inline.contexts[context as usize].link.clone();
        inline.contexts.push(TextContext {
            styles,
            parent: Some(context),
            link,
        });
        let index = (inline.contexts.len() - 1) as u32;
        inline.by_span.insert((context, name.to_string()), index);
        Ok(index)
    }

    fn link_context(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<u32, String> {
        if inline.contexts.len() >= MAX_PARAGRAPH_CONTEXTS {
            return Ok(context);
        }
        let href = node.attribute((NS_XLINK, "href")).unwrap_or("");
        let target = match href.strip_prefix('#') {
            Some(anchor) => (None, Some(anchor.to_string())),
            None if !href.is_empty() => (Some(href.to_string()), None),
            None => (None, None),
        };
        let styles: Rc<[StyleRef<'a>]> = match node
            .attribute((NS_TEXT, "style-name"))
            .and_then(|name| self.find_style(Family::Text, name))
        {
            Some(style) => self.style_chain(style)?,
            None => Rc::from(Vec::new()),
        };
        inline.contexts.push(TextContext {
            styles,
            parent: Some(context),
            link: Some(Rc::new(target)),
        });
        Ok((inline.contexts.len() - 1) as u32)
    }

    fn field(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        kind: &'static str,
        instruction: &'static str,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<(), String> {
        let fallback = text_content(node)?;
        self.admit_chars(1, inline)?;
        let index = inline.items.len() as u32;
        inline.items.push(Item::Field {
            kind,
            instruction,
            fallback,
        });
        inline.tokens.push(Token::Item(index, context));
        Ok(())
    }

    /// `text:ruby` (§6.4): the base text with its annotation.
    fn ruby(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<(), String> {
        let text_of = |name: &str| -> Result<String, String> {
            match node.kids().find(|n| xml::is(*n, NS_TEXT, name)) {
                Some(child) => text_content(child),
                None => Ok(String::new()),
            }
        };
        let base = text_of("ruby-base")?;
        let text = text_of("ruby-text")?;
        if base.is_empty() {
            return Ok(());
        }
        // The ruby text's own style sets its size; LibreOffice's default
        // ruby text is half the base size.
        // (The style and its parents, §16.2.)
        let ruby_size = match node
            .kids()
            .find(|n| xml::is(*n, NS_TEXT, "ruby-text"))
            .and_then(|n| n.attribute((NS_TEXT, "style-name")))
            .and_then(|name| self.find_style(Family::Text, name))
        {
            Some(style) => {
                let cascade = Cascade::build(self.style_chain(style)?.to_vec(), TEXT)?;
                // A percentage refers to the parent style's size (§20.190);
                // past the style's own parents the search continues in the
                // text the ruby annotates (§16.2), so to the base text's.
                // (Only where the chain sets a size at all: an unstyled
                // ruby text is LibreOffice's half size.)
                (cascade.get(NS_FO, "font-size").is_some()
                    || cascade.get(NS_STYLE, "font-size-rel").is_some())
                .then(|| size_terms(&cascade, ""))
            }
            None => None,
        };
        self.admit_chars(base.chars().count(), inline)?;
        let index = inline.items.len() as u32;
        inline.items.push(Item::Ruby {
            base,
            text,
            ruby_size,
        });
        inline.tokens.push(Token::Item(index, context));
        Ok(())
    }

    /// `text:note` (§6.3): the citation in the text and the note body as a
    /// DOCX footnote or endnote.
    fn note(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        context: u32,
        inline: &mut Inline<'a>,
    ) -> Result<(), String> {
        let kind = match node.attribute((NS_TEXT, "note-class")) {
            Some("endnote") => "endnote",
            _ => "footnote",
        };
        let citation = node.kids().find(|n| xml::is(*n, NS_TEXT, "note-citation"));
        let label = citation
            .and_then(|c| c.attribute((NS_TEXT, "label")))
            .map(str::to_string);
        self.next_note += 1;
        let id = self.next_note.to_string();
        let width = self.text_width;
        let mut content = match node.kids().find(|n| xml::is(*n, NS_TEXT, "note-body")) {
            Some(body) => {
                let saved_scope = self.scope;
                let result = self.contained_story(body, width);
                self.scope = saved_scope;
                result?
            }
            None => Vec::new(),
        };
        // The note's number before its first paragraph's text, formatted by
        // the configuration's citation style (§16.29.3
        // `text:citation-style-name`), followed by a tab to the hanging
        // indent where the paragraph has one (LibreOffice places the body
        // text at the indent).
        let style_name = self.notes_config.citation_style(kind);
        let style = Some(style_name.to_string());
        if let Some(BodyElement::Paragraph(first)) = content.first_mut() {
            let mut runs = Vec::new();
            let mark = self.note_mark_run(kind, style.as_deref(), first, label.as_deref())?;
            runs.extend(mark);
            if first.indent_first < 0.0 {
                let mut tab = match runs.last() {
                    Some(DocRun::Text(run)) => run.as_ref().clone(),
                    _ => TextRun::default(),
                };
                tab.text = "\t".into();
                tab.note_ref = None;
                tab.vert_align = None;
                tab.position = None;
                runs.push(DocRun::Text(Box::new(tab)));
            }
            first.runs.splice(0..0, runs);
        }
        let note = DocxNote {
            id: id.clone(),
            content,
        };
        if kind == "endnote" {
            self.endnotes.push(note);
        } else {
            self.footnotes.push(note);
        }
        // The citation in the text flow takes the configuration's body
        // citation style (§19.773) over the enclosing formatting.
        let body_style = self.notes_config.body_citation_style(kind);
        let citation_context = match self.find_common_style(Family::Text, body_style) {
            Some(style) if inline.contexts.len() < MAX_PARAGRAPH_CONTEXTS => {
                let styles = self.style_chain(style)?;
                let link = inline.contexts[context as usize].link.clone();
                inline.contexts.push(TextContext {
                    styles,
                    parent: Some(context),
                    link,
                });
                (inline.contexts.len() - 1) as u32
            }
            _ => context,
        };
        let index = inline.items.len() as u32;
        inline.items.push(Item::NoteReference { kind, id, label });
        self.admit_chars(1, inline)?;
        inline.tokens.push(Token::Item(index, citation_context));
        Ok(())
    }

    /// The number run at the start of a note's text.
    fn note_mark_run(
        &mut self,
        kind: &'static str,
        style: Option<&str>,
        paragraph: &DocParagraph,
        label: Option<&str>,
    ) -> Result<Vec<DocRun>, String> {
        let mut base = TextRun {
            font_size: paragraph.default_font_size.unwrap_or(10.0),
            font_family: paragraph.default_font_family.clone(),
            font_family_high_ansi: paragraph.default_font_family.clone(),
            font_family_east_asia: paragraph.default_font_family_east_asia.clone(),
            font_family_cs: paragraph.default_font_family.clone(),
            ..TextRun::default()
        };
        if let Some(style) = style.and_then(|name| self.find_common_style(Family::Text, name)) {
            // The citation style's search continues in the note paragraph
            // (§16.2), which is already built: what the style and its
            // parents leave unset, size and face, is the paragraph's.
            let chain = self.style_chain(style)?;
            let sets_face = Cascade::build(chain.to_vec(), TEXT)?
                .keys()
                .iter()
                .any(|(_, name)| name.starts_with("font-name") || name.starts_with("font-family"));
            let inline = Inline::new(chain);
            let mut properties = self.run_properties(&inline.contexts, 0, ScriptClass::Latin)?;
            // A relative size (the citation style's 58% superscript) refers
            // to the note paragraph's own size.
            if let Some(size) = paragraph.default_font_size {
                properties.size = properties
                    .size_relative
                    .map_or(properties.size, |f| f * size);
            }
            let styled = properties.text_run();
            base = if sets_face {
                styled
            } else {
                TextRun {
                    font_family: base.font_family,
                    font_family_high_ansi: base.font_family_high_ansi,
                    font_family_east_asia: base.font_family_east_asia,
                    font_family_cs: base.font_family_cs,
                    ..styled
                }
            };
        }
        // The mark run, and the custom label run when there is one, each own
        // a copy of the base properties.
        let mut runs = Vec::new();
        let mut mark = base.clone();
        mark.text = String::new();
        mark.note_ref = Some(NoteRef {
            kind: kind.into(),
            id: String::new(),
            custom_mark_follows: label.is_some(),
        });
        runs.push(DocRun::Text(Box::new(mark)));
        if let Some(label) = label {
            let mut custom = base;
            custom.text = label.to_string();
            runs.push(DocRun::Text(Box::new(custom)));
        }
        Ok(runs)
    }

    pub(super) fn run_properties(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
        class: ScriptClass,
    ) -> Result<RunProperties, String> {
        let styles = self.search_order(contexts, context)?;
        let cascade = Cascade::build(styles, TEXT)?;
        let (properties, face) = RunProperties::resolve(&cascade, class, self.fonts);
        self.used_faces.borrow_mut().insert(face);
        Ok(properties)
    }

    /// The §16.2 search order of one context, nearest style first: the
    /// span chains, the paragraph chain, then the text default style.
    fn search_order(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<Vec<StyleRef<'a>>, String> {
        let mut styles = Vec::new();
        let mut current = Some(context);
        while let Some(index) = current {
            let entry = &contexts[index as usize];
            styles.extend(entry.styles.iter().copied());
            if styles.len() > MAX_SEARCH_ORDER {
                return Err(resource("style search order"));
            }
            current = entry.parent;
        }
        if let Some(default) = self.sheet.default_style(Family::Text) {
            styles.push(default);
        }
        Ok(styles)
    }

    /// Script classes (§20.358 Table 22) and run emission.
    pub(super) fn emit_runs(
        &mut self,
        inline: &mut Inline<'a>,
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        let tokens = collapse(std::mem::take(&mut inline.tokens));
        let contexts = &inline.contexts;
        let justified = paragraph.alignment == "both";
        let balanced = if self.settings.balance {
            self.balanced_spaces(&tokens, contexts, justified)?
        } else {
            Vec::new()
        };
        let weak = weak_classes(&tokens);
        let mut cache: HashMap<(u32, ScriptClass), RunProperties> = HashMap::new();
        // The last emitted text run's key and direction, for merging.
        let mut current: Option<(u32, ScriptClass, bool)> = None;
        let mut previous_class = ScriptClass::Latin;
        let mut previous_rtl = false;
        // Whether the previous character starts a word (capitalize).
        let mut word_start = true;
        for (index, &token) in tokens.iter().enumerate() {
            let (c, context) = match token {
                Token::LineBreak => {
                    paragraph.runs.push(DocRun::Break {
                        break_type: BreakType::Line,
                    });
                    current = None;
                    word_start = true;
                    continue;
                }
                Token::Boundary(_) => continue,
                Token::Item(item, context) => {
                    let key = (context, previous_class);
                    if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(key) {
                        entry.insert(self.run_properties(contexts, context, key.1)?);
                    }
                    let properties = &cache[&key];
                    let run = self.item_run(
                        &mut inline.items[item as usize],
                        properties,
                        contexts,
                        context,
                    )?;
                    paragraph.runs.extend(run);
                    current = None;
                    continue;
                }
                Token::Tab(context) => ('\t', context),
                Token::Char(c, context) => (c, context),
                Token::Space(context) => (' ', context),
            };
            let class = table_22(c)
                .or_else(|| weak.get(index).copied().flatten())
                .unwrap_or(previous_class);
            let rtl = class == ScriptClass::Complex
                && (strong_rtl(c) || (!c.is_alphanumeric() && previous_rtl));
            previous_class = class;
            previous_rtl = rtl;
            let key = (context, class);
            if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(key) {
                entry.insert(self.run_properties(contexts, context, class)?);
            }
            let properties = &cache[&key];
            // §20.367 text:display="none": hidden text.
            if properties.hidden {
                continue;
            }
            // Text transformations the DOCX run cannot carry (§20.227):
            // lower case and title case are applied to the characters.
            let transform = match properties.transform {
                Transform::Lowercase => Transform::Lowercase,
                Transform::Capitalize if word_start => Transform::Capitalize,
                _ => Transform::None,
            };
            let push = |text: &mut String| match transform {
                Transform::Lowercase => text.extend(c.to_lowercase()),
                Transform::Capitalize => text.extend(c.to_uppercase()),
                Transform::None => text.push(c),
            };
            word_start = c.is_whitespace() || c == '\t';
            if c == '\u{ad}' {
                // §6.1.6 soft hyphen: an optional hyphen run.
                budget::tick(1)?;
                let mut run = properties.text_run();
                run.optional_hyphen = true;
                self.apply_context(&mut run, contexts, context)?;
                paragraph.runs.push(DocRun::Text(Box::new(run)));
                current = None;
                continue;
            }
            let advance = balanced.get(index).copied().flatten();
            let continues = match (current, paragraph.runs.last_mut()) {
                (Some(last), Some(DocRun::Text(run)))
                    if last == (context, class, rtl) && run.space_advance_pt == advance =>
                {
                    push(&mut run.text);
                    true
                }
                _ => false,
            };
            if !continues {
                budget::tick(1)?;
                let mut run = properties.text_run();
                push(&mut run.text);
                run.space_advance_pt = advance;
                if rtl {
                    run.rtl = Some(true);
                }
                self.apply_context(&mut run, contexts, context)?;
                paragraph.runs.push(DocRun::Text(Box::new(run)));
            }
            current = Some((context, class, rtl));
        }
        for name in std::mem::take(&mut inline.bookmarks) {
            paragraph.bookmarks.push(name);
        }
        Ok(())
    }

    /// The hyperlink of a context, on a run (each run owns its copy of the
    /// target).
    fn apply_context(
        &mut self,
        run: &mut TextRun,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<(), String> {
        if let Some(link) = &contexts[context as usize].link {
            run.is_link = true;
            run.hyperlink = link.0.clone();
            run.hyperlink_anchor = link.1.clone();
        }
        Ok(())
    }

    /// The runs of one item. Each item is emitted once, so its prebuilt run
    /// moves out of the item (no second copy).
    fn item_run(
        &mut self,
        item: &mut Item,
        properties: &RunProperties,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<Vec<DocRun>, String> {
        budget::tick(1)?;
        Ok(match item {
            Item::Run(run) => {
                let mut run = std::mem::replace(
                    run,
                    DocRun::Break {
                        break_type: BreakType::Line,
                    },
                );
                match &mut run {
                    DocRun::Image(image) if !image.anchor => {
                        image.inline_object_font_facts =
                            Some(Box::new(font_facts_of(&properties.text_run())));
                    }
                    DocRun::Shape(shape) if shape.inline => {
                        shape.inline_object_font_facts =
                            Some(Box::new(font_facts_of(&properties.text_run())));
                    }
                    _ => {}
                }
                with_anchor_host(run, &properties.text_run())
            }
            Item::Field {
                kind,
                instruction,
                fallback,
            } => {
                let text = properties.text_run();
                vec![DocRun::Field(Box::new(FieldRun {
                    field_type: (*kind).into(),
                    instruction: (*instruction).into(),
                    fallback_text: fallback.clone(),
                    bold: text.bold,
                    italic: text.italic,
                    underline: text.underline,
                    strikethrough: text.strikethrough,
                    font_size: text.font_size,
                    color: text.color.clone(),
                    font_family: text.font_family.clone(),
                    font_family_high_ansi: text.font_family_high_ansi.clone(),
                    font_family_east_asia: text.font_family_east_asia.clone(),
                    font_family_cs: text.font_family_cs.clone(),
                    font_size_cs: text.font_size_cs,
                    bold_cs: text.bold_cs,
                    italic_cs: text.italic_cs,
                    background: text.background.clone(),
                    vert_align: text.vert_align.clone(),
                    all_caps: text.all_caps,
                    small_caps: text.small_caps,
                    double_strikethrough: text.double_strikethrough,
                    ..FieldRun::default()
                }))]
            }
            Item::NoteReference { kind, id, label } => {
                let mut run = properties.text_run();
                run.text = label.clone().unwrap_or_else(|| id.clone());
                run.note_ref = Some(NoteRef {
                    kind: (*kind).into(),
                    id: id.clone(),
                    custom_mark_follows: label.is_some(),
                });
                if label.is_some() {
                    let mut custom = properties.text_run();
                    custom.text = label.clone().unwrap_or_default();
                    run.text = String::new();
                    vec![DocRun::Text(Box::new(run)), DocRun::Text(Box::new(custom))]
                } else {
                    vec![DocRun::Text(Box::new(run))]
                }
            }
            Item::Ruby {
                base,
                text,
                ruby_size,
            } => {
                let mut run = properties.text_run();
                run.text = base.clone();
                run.ruby = Some(RubyAnnotation {
                    text: text.clone(),
                    font_size_pt: match *ruby_size {
                        Some((base, factor, add)) => {
                            (base.unwrap_or(properties.size) * factor + add).max(0.5)
                        }
                        None => properties.size / 2.0,
                    },
                    hps_raise_pt: None,
                    typography: None,
                });
                self.apply_context(&mut run, contexts, context)?;
                vec![DocRun::Text(Box::new(run))]
            }
        })
    }

    /// BalanceSpacesAndIdeographicSpaces (LibreOffice `sw::Justify::
    /// BalanceCjkSpaces`, measured): a U+0020 that is the first or last
    /// character of a text portion, or next to another U+0020, advances half
    /// its Asian font size. Portions end at the paragraph's ends, line
    /// breaks, soft hyphens, bookmarks and reference marks, objects, and
    /// where the span formatting changes. Returns each token's advance in
    /// points (`None` for other characters).
    fn balanced_spaces(
        &mut self,
        tokens: &[Token],
        contexts: &[TextContext<'a>],
        _justified: bool,
    ) -> Result<Vec<Option<f64>>, String> {
        let space =
            |token: Option<&Token>| matches!(token, Some(Token::Char(' ', _) | Token::Space(_)));
        let context = |token: &Token| match *token {
            Token::Char(_, context) | Token::Space(context) | Token::Tab(context) => Some(context),
            _ => None,
        };
        let mut advances: HashMap<u32, f64> = HashMap::new();
        let mut out = vec![None; tokens.len()];
        for (index, token) in tokens.iter().enumerate() {
            if !space(Some(token)) {
                continue;
            }
            let own = context(token).expect("a space has a context");
            let previous = index.checked_sub(1).map(|i| &tokens[i]);
            let next = tokens.get(index + 1);
            let mut balanced = space(previous) || space(next);
            for neighbour in [previous, next] {
                if balanced {
                    break;
                }
                balanced = match neighbour {
                    None => true,
                    Some(Token::LineBreak | Token::Boundary(true) | Token::Item(..)) => true,
                    Some(Token::Boundary(false)) => false,
                    Some(Token::Char(c, _)) if portion_character(*c) => true,
                    Some(other) => context(other).is_some_and(|other| {
                        other != own
                            && !Rc::ptr_eq(
                                &contexts[other as usize].styles,
                                &contexts[own as usize].styles,
                            )
                    }),
                };
            }
            if !balanced {
                continue;
            }
            let advance = match advances.get(&own) {
                Some(advance) => *advance,
                None => {
                    let asian = self.run_properties(contexts, own, ScriptClass::Asian)?;
                    // LibreOffice keeps font heights in twips.
                    let advance = (asian.size * 20.0).round() / 40.0;
                    advances.insert(own, advance);
                    advance
                }
            };
            out[index] = Some(advance);
        }
        Ok(out)
    }
}

/// A floating drawing preceded by the zero-advance metrics of the character
/// that carries its anchor (the DOCX anchor host), whose occurrence id ties
/// the two together.
/// Whether `text:hidden-paragraph` fields (§7.7.11) hide this paragraph:
/// one whose condition last evaluated true (`text:is-hidden="true"`,
/// §19.821) hides it, unless the paragraph holds another whose condition
/// evaluated false. The fields sit in the paragraph's inline content (spans,
/// links, meta), not in notes or drawings it carries.
pub(super) fn hidden_by_field(paragraph: roxmltree::Node<'_, '_>) -> Result<bool, String> {
    /// (a field hides, a field shows) among `node`'s inline content.
    fn walk(
        node: roxmltree::Node<'_, '_>,
        depth: usize,
        found: &mut (bool, bool),
    ) -> Result<(), String> {
        if depth > MAX_INLINE_DEPTH {
            return Ok(());
        }
        for child in node.kids().filter(|n| n.is_element()) {
            budget::tick(1)?;
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_TEXT), "hidden-paragraph") => {
                    match child.attribute((NS_TEXT, "is-hidden")) {
                        Some("true") => found.0 = true,
                        Some("false") => found.1 = true,
                        _ => {}
                    }
                }
                (Some(NS_TEXT), "note") => {}
                (Some(NS_TEXT), _) | (Some(NS_LOEXT), _) => walk(child, depth + 1, found)?,
                _ => {}
            }
        }
        Ok(())
    }
    let mut found = (false, false);
    walk(paragraph, 0, &mut found)?;
    Ok(found.0 && !found.1)
}

pub(super) fn with_anchor_host(run: DocRun, host: &TextRun) -> Vec<DocRun> {
    let occurrence = match &run {
        DocRun::Image(image) if image.anchor => image.anchor_acquisition.as_ref(),
        DocRun::Shape(shape) if !shape.inline => shape.anchor_acquisition.as_ref(),
        _ => None,
    }
    .map(|facts| facts.occurrence_id.clone());
    let Some(occurrence) = occurrence else {
        return vec![run];
    };
    vec![
        DocRun::AnchorHost(docx_model::AnchorHostMetrics {
            font_size: host.font_size,
            font_family: host.font_family.clone(),
            font_family_east_asia: host.font_family_east_asia.clone(),
            bold: host.bold,
            italic: host.italic,
            anchor_occurrence_id: Some(occurrence),
            list_label: None,
        }),
        run,
    ]
}

/// Characters LibreOffice sets in text portions of their own (measured):
/// soft hyphen, no-break space, non-breaking hyphen, zero-width space and
/// word joiner.
fn portion_character(c: char) -> bool {
    matches!(
        c,
        '\u{ad}' | '\u{a0}' | '\u{2011}' | '\u{200b}' | '\u{2060}'
    )
}

/// Strong right-to-left characters (Hebrew, Arabic, Syriac, Thaana, NKo and
/// their presentation forms).
fn strong_rtl(c: char) -> bool {
    matches!(u32::from(c),
        0x0590..=0x08FF | 0xFB1D..=0xFDFF | 0xFE70..=0xFEFF | 0x10800..=0x10FFF
        | 0x1E800..=0x1EFFF)
}

/// LibreOffice's class for weak characters outside Table 22 (measured;
/// user-approved compatibility rule 2026-10-31): the class of the preceding
/// character, or at paragraph start of the following one.
fn weak_classes(tokens: &[Token]) -> Vec<Option<ScriptClass>> {
    let character = |token: &Token| match *token {
        Token::Char(c, _) => Some(c),
        Token::Space(_) | Token::Tab(_) => Some(' '),
        _ => None,
    };
    let mut classes: Vec<Option<ScriptClass>> = Vec::with_capacity(tokens.len());
    let mut previous: Option<ScriptClass> = None;
    for token in tokens {
        let class = match character(token) {
            Some(c) => match table_22(c) {
                Some(class) => Some(class),
                None if is_weak_unmapped(c) || c == ' ' => previous,
                None => previous,
            },
            None => None,
        };
        if class.is_some() {
            previous = class;
        }
        classes.push(class);
    }
    if let Some(first) = classes.iter().position(Option::is_some) {
        let following = classes[first];
        for slot in classes.iter_mut().take(first) {
            *slot = following;
        }
    }
    classes
}

/// §6.1.2 white-space collapsing over a paragraph's character data: tab, CR
/// and LF become spaces, leading and trailing spaces are removed and runs of
/// spaces collapse to one. `text:s`, `text:tab`, `text:line-break` and
/// objects keep their characters and are not collapsed.
fn collapse(tokens: Vec<Token>) -> Vec<Token> {
    let mut out: Vec<Token> = Vec::with_capacity(tokens.len());
    let mut pending: Option<Token> = None;
    let mut boundaries: Vec<Token> = Vec::new();
    let mut started = false;
    for token in tokens {
        if let Token::Boundary(_) = token {
            if pending.is_some() {
                boundaries.push(token);
            } else {
                out.push(token);
            }
            continue;
        }
        match token {
            Token::Char(' ' | '\t' | '\r' | '\n', context) => {
                if pending.is_none() && started {
                    pending = Some(Token::Char(' ', context));
                }
            }
            Token::Char(c, context) => {
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                out.append(&mut boundaries);
                started = true;
                out.push(Token::Char(c, context));
            }
            element => {
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                out.append(&mut boundaries);
                started = true;
                out.push(element);
            }
        }
    }
    out
}

// ---------------------------------------------------------------------------
// Character properties

#[derive(Clone, Copy, PartialEq, Debug)]
pub(super) enum Transform {
    None,
    Lowercase,
    Capitalize,
}

/// Effective text properties of one script class.
#[derive(Clone, PartialEq, Debug)]
pub(super) struct RunProperties {
    pub family: Atom,
    pub size: f64,
    /// The nearest relative size factor when the size is a percentage of an
    /// unknown base (for note citations).
    pub size_relative: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<&'static str>,
    pub underline_color: Option<String>,
    pub strike: Option<&'static str>,
    pub color: Option<String>,
    pub color_auto: bool,
    pub background: Option<String>,
    pub letter_spacing: Option<f64>,
    pub scale: Option<f64>,
    pub position: Option<f64>,
    pub caps: bool,
    pub small_caps: bool,
    pub transform: Transform,
    pub kerning: Option<f64>,
    pub hidden: bool,
    pub emphasis: Option<&'static str>,
    pub combine: bool,
    pub vertical: bool,
    pub border: Option<(String, Option<String>, u64)>,
    pub lang_east_asia: Option<Atom>,
    pub lang_bidi: Option<Atom>,
    pub complex: bool,
}

impl RunProperties {
    fn resolve(cascade: &Cascade<'_>, class: ScriptClass, fonts: &FontFaces) -> (Self, Face) {
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
                cascade
                    .get(NS_STYLE, &format!("{name}{suffix}"))
                    .or_else(|| cascade.get(NS_FO, name))
            }
        };
        let face = resolve_family(cascade, suffix, fonts)
            .or_else(|| resolve_family(cascade, "", fonts))
            .unwrap_or_else(|| Face {
                family: fonts.atom("Liberation Serif"),
                generic: Some("roman"),
                pitch: None,
                charset: None,
            });
        let family = face.family.clone();
        let (mut size, size_relative) = resolve_size(cascade, suffix);
        let bold = match scripted("font-weight") {
            Some("bold" | "bolder") => true,
            Some(value) => value.parse::<u32>().is_ok_and(|weight| weight >= 600),
            None => false,
        };
        let italic = matches!(scripted("font-style"), Some("italic" | "oblique"));
        let (underline, underline_color) = underline(cascade);
        let strike = match cascade.get(NS_STYLE, "text-line-through-style") {
            None | Some("none") => {
                // A line-through text (§20.367) also strikes the characters.
                if cascade
                    .get(NS_STYLE, "text-line-through-text")
                    .is_some_and(|t| !t.is_empty())
                {
                    Some("single")
                } else {
                    None
                }
            }
            Some(_) => match cascade.get(NS_STYLE, "text-line-through-type") {
                Some("double") => Some("double"),
                Some("none") => None,
                _ => Some("single"),
            },
        };
        let background = match cascade.get(NS_FO, "background-color") {
            None | Some("transparent") => None,
            Some(value) => hex_color(value).ok(),
        };
        let (mut color_auto, mut color) = font_color(cascade);
        // Character transparency (LibreOffice `loext:opacity`) over a white
        // page: the colour blended with white.
        if let (Some(value), Some(c)) = (cascade.get(NS_LOEXT, "opacity"), color.as_ref()) {
            if let Ok(opacity) = percent(value) {
                if (0.0..1.0).contains(&opacity) {
                    color = Some(blend_white(c, opacity));
                }
            }
        }
        // §20.395 automatic (or unset) colour over the run's own shading.
        if let (None, Some(background)) = (&color, &background) {
            color = Some(libreoffice_automatic_color(background).to_string());
            color_auto = false;
        }
        let letter_spacing = cascade
            .get(NS_FO, "letter-spacing")
            .filter(|v| *v != "normal")
            .and_then(|v| length_pt(v).ok())
            .filter(|v| *v != 0.0);
        let scale = cascade
            .get(NS_STYLE, "text-scale")
            .and_then(|v| percent(v).ok())
            .filter(|v| *v > 0.0 && (*v - 1.0).abs() > 1e-9);
        // §20.374 style:text-position: a raise (a percentage of the font
        // size) and an optional relative font size. The keywords `super` and
        // `sub` are LibreOffice's automatic escapement (SwSubFont::CalcEsc):
        // a superscript's ascent ends at the full-size ascent, a
        // subscript's descent at the full-size descent. The reader has no
        // font metrics, so the em is split at the nominal 0.8 / 0.2 of a
        // font's ascent and descent.
        let mut position = None;
        if let Some(value) = cascade.get(NS_STYLE, "text-position") {
            let mut parts = value.split_whitespace();
            let raise = parts.next();
            let relative = parts
                .next()
                .and_then(|v| percent(v).ok())
                .filter(|r| *r > 0.0)
                .unwrap_or(1.0);
            let raise = match raise {
                Some("super") => Some(0.8 * (1.0 - relative.min(1.0)).max(0.0)),
                Some("sub") => Some(-0.2 * (1.0 - relative.min(1.0)).max(0.0)),
                Some(v) => percent(v).ok(),
                None => None,
            };
            if let Some(raise) = raise.filter(|r| *r != 0.0) {
                position = Some(raise * size);
            }
            if (relative - 1.0).abs() > 1e-9 {
                size *= relative;
            }
        }
        let transform_value = cascade.get(NS_FO, "text-transform");
        let caps = transform_value == Some("uppercase");
        let transform = match transform_value {
            Some("lowercase") => Transform::Lowercase,
            Some("capitalize") => Transform::Capitalize,
            _ => Transform::None,
        };
        let small_caps = cascade.get(NS_FO, "font-variant") == Some("small-caps");
        // §20.315 pair kerning; LibreOffice kerns text whose styles never set
        // the attribute (measured).
        let kerning = match cascade.get(NS_STYLE, "letter-kerning") {
            Some("false") => None,
            _ => Some(KERN_EVERY_SIZE),
        };
        let hidden = cascade.get(NS_TEXT, "display") == Some("none");
        // §20.360 style:text-emphasize.
        let emphasis = cascade.get(NS_STYLE, "text-emphasize").and_then(|v| {
            let mut parts = v.split_whitespace();
            let mark = parts.next()?;
            let below = parts.next() == Some("below");
            match mark {
                "none" => None,
                "dot" | "disc" if below => Some("underDot"),
                "dot" | "disc" => Some("dot"),
                "circle" => Some("circle"),
                "accent" => Some("comma"),
                _ => None,
            }
        });
        let combine = cascade.get(NS_STYLE, "text-combine") == Some("lines");
        let vertical = matches!(
            cascade.get(NS_STYLE, "text-rotation-angle"),
            Some("90" | "270" | "90deg" | "270deg")
        );
        let border = cascade
            .get(NS_LOEXT, "border")
            .and_then(|value| parse_border(value, None))
            .map(|spec| {
                (
                    spec.style.to_string(),
                    spec.color,
                    (spec.width * 1000.0).round() as u64,
                )
            });
        let language = |language: &str, country: &str| -> Option<Atom> {
            let language = cascade.get(NS_STYLE, language).filter(|v| *v != "none")?;
            Some(fonts.atom(
                &match cascade.get(NS_STYLE, country).filter(|v| *v != "none") {
                    Some(country) => format!(
                        "{}-{}",
                        language.to_ascii_lowercase(),
                        country.to_ascii_uppercase()
                    ),
                    None => language.to_ascii_lowercase(),
                },
            ))
        };
        if size < 0.5 {
            size = 0.5;
        }
        (
            Self {
                family,
                size,
                size_relative,
                bold,
                italic,
                underline,
                underline_color,
                strike,
                color,
                color_auto,
                background,
                letter_spacing,
                scale,
                position,
                caps,
                small_caps,
                transform,
                kerning,
                hidden,
                emphasis,
                combine,
                vertical,
                border,
                lang_east_asia: language("language-asian", "country-asian"),
                lang_bidi: language("language-complex", "country-complex"),
                complex: class == ScriptClass::Complex,
            },
            face,
        )
    }

    /// Native size of one emitted run, excluding its text and link: the run
    /// shells and every string `text_run` copies out of the properties.
    pub fn text_run(&self) -> TextRun {
        let family = Some(self.family.to_string());
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
            underline_color: self.underline_color.clone(),
            strikethrough: self.strike == Some("single"),
            double_strikethrough: self.strike == Some("double"),
            font_size: self.size,
            font_size_cs: Some(self.size),
            color: self.color.clone(),
            color_auto: self.color_auto,
            font_family: family.clone(),
            font_family_high_ansi: family.clone(),
            font_family_east_asia: family.clone(),
            font_family_cs: family,
            cs: self.complex.then_some(true),
            background: self.background.clone(),
            char_spacing: self.letter_spacing,
            char_scale: self.scale,
            position: self.position,
            all_caps: self.caps,
            small_caps: self.small_caps,
            kerning: self.kerning,
            emphasis_mark: self.emphasis.map(str::to_string),
            east_asian_combine: self.combine.then_some(true),
            east_asian_vert: self.vertical.then_some(true),
            border: self.border.as_ref().map(|(style, color, width)| RunBorder {
                style: style.clone(),
                color: color.clone(),
                width: *width as f64 / 1000.0,
                space: 0.0,
            }),
            lang_east_asia: self.lang_east_asia.as_deref().map(str::to_string),
            lang_bidi: self.lang_bidi.as_deref().map(str::to_string),
            ..TextRun::default()
        }
    }

    pub fn font_facts(&self) -> RunFontFacts {
        let family = Some(self.family.to_string());
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

/// The font size a cascade sets for a script class (§20.190 fo:font-size,
/// §20.286 style:font-size-rel), as `base * factor + add`: the nearest
/// absolute size (`None` when the search order has none), the product of
/// the percentages set nearer, and the relative sizes set nearer, each
/// scaled by the percentages nearer still.
fn size_terms(cascade: &Cascade<'_>, suffix: &str) -> (Option<f64>, f64, f64) {
    let (size_key, rel_key) = if suffix.is_empty() {
        (
            (NS_FO.to_string(), "font-size".to_string()),
            (NS_STYLE.to_string(), "font-size-rel".to_string()),
        )
    } else {
        (
            (NS_STYLE.to_string(), format!("font-size{suffix}")),
            (NS_STYLE.to_string(), format!("font-size-rel{suffix}")),
        )
    };
    let mut factor = 1.0;
    let mut add = 0.0;
    for style in &cascade.styles {
        let Some(group) = style.group(cascade.group) else {
            continue;
        };
        if let Some(value) = group.attributes.get(&size_key) {
            if value.ends_with('%') {
                factor *= percent(value).unwrap_or(1.0);
                continue;
            }
            if let Ok(size) = length_pt(value) {
                return (Some(size), factor, add);
            }
        } else if let Some(value) = group.attributes.get(&rel_key) {
            add += length_pt(value).unwrap_or(0.0) * factor;
        }
    }
    (None, factor, add)
}

/// The font size of a script class. Returns the size and, when no absolute
/// size exists, the relative factor (for a caller that knows what the size
/// is relative to).
fn resolve_size(cascade: &Cascade<'_>, suffix: &str) -> (f64, Option<f64>) {
    match size_terms(cascade, suffix) {
        (Some(size), factor, add) => ((size * factor + add).max(0.5), None),
        // A script without its own size takes the western one.
        (None, ..) if !suffix.is_empty() => resolve_size(cascade, ""),
        // ODF defines no initial font size; LibreOffice's is 12 pt.
        (None, factor, add) => ((12.0 * factor + add).max(0.5), Some(factor)),
    }
}

/// The western font size of a text-properties cascade.
pub(super) fn paragraph_font_size(cascade: &Cascade<'_>) -> f64 {
    resolve_size(cascade, "").0
}

/// §20.389-§20.394 underline: style, type, width and colour mapped to the
/// DOCX underline vocabulary (ST_Underline, §17.18.99).
fn underline(cascade: &Cascade<'_>) -> (Option<&'static str>, Option<String>) {
    let style = cascade
        .get(NS_STYLE, "text-underline-style")
        .unwrap_or("none");
    let kind = cascade
        .get(NS_STYLE, "text-underline-type")
        .unwrap_or("single");
    if style == "none" || kind == "none" {
        return (None, None);
    }
    let bold = matches!(
        cascade.get(NS_STYLE, "text-underline-width"),
        Some("bold" | "thick" | "medium")
    );
    let words = cascade.get(NS_STYLE, "text-underline-mode") == Some("skip-white-space");
    let double = kind == "double";
    let value = match (style, double, bold) {
        ("wave", true, _) => "wavyDouble",
        ("wave", _, true) => "wavyHeavy",
        ("wave", _, _) => "wave",
        (_, true, _) => "double",
        ("dotted", _, true) => "dottedHeavy",
        ("dotted", _, _) => "dotted",
        ("dash", _, true) => "dashedHeavy",
        ("dash", _, _) => "dash",
        ("long-dash", _, true) => "dashLongHeavy",
        ("long-dash", _, _) => "dashLong",
        ("dot-dash", _, true) => "dashDotHeavy",
        ("dot-dash", _, _) => "dotDash",
        ("dot-dot-dash", _, true) => "dashDotDotHeavy",
        ("dot-dot-dash", _, _) => "dotDotDash",
        (_, _, true) => "thick",
        _ if words => "words",
        _ => "single",
    };
    let color = cascade
        .get(NS_STYLE, "text-underline-color")
        .filter(|c| *c != "font-color")
        .and_then(|c| hex_color(c).ok());
    (Some(value), color)
}

/// The face of a script slot: the nearest style that sets either the face
/// reference (`style:font-name*`) or a direct family (`fo:font-family`,
/// `style:font-family-asian/-complex`); within one style the face reference
/// wins.
fn resolve_family(cascade: &Cascade<'_>, suffix: &str, fonts: &FontFaces) -> Option<Face> {
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
            if let Some(face) = fonts.faces.get(name) {
                return Some(face.clone());
            }
            // A missing declaration: the name is the family (LibreOffice).
            return Some(Face {
                family: fonts.atom(name),
                generic: None,
                pitch: None,
                charset: None,
            });
        }
        if let Some(family) = group.attributes.get(&family_key) {
            let attribute = |name: &str| {
                group
                    .attributes
                    .get(&(NS_STYLE.to_string(), format!("{name}{suffix}")))
                    .map(String::as_str)
            };
            return Some(Face {
                family: fonts.atom(&first_family(family)),
                generic: generic_class(attribute("font-family-generic")),
                pitch: pitch(attribute("font-pitch")),
                charset: charset(attribute("font-charset")),
            });
        }
    }
    None
}

/// The effective font colour: `(automatic, explicit colour)`. LibreOffice
/// lets the nearest style that sets `fo:color` or
/// `style:use-window-font-color="true"` decide (user-approved compatibility
/// rule, 2026-10-31).
fn font_color(cascade: &Cascade<'_>) -> (bool, Option<String>) {
    let window = (NS_STYLE.to_string(), "use-window-font-color".to_string());
    let explicit = (NS_FO.to_string(), "color".to_string());
    for style in &cascade.styles {
        let Some(properties) = style.group(cascade.group) else {
            continue;
        };
        if properties.attributes.get(&window).map(String::as_str) == Some("true") {
            return (true, None);
        }
        if let Some(color) = properties.attributes.get(&explicit) {
            return (false, hex_color(color).ok());
        }
    }
    (false, None)
}

fn blend_white(color: &str, opacity: f64) -> String {
    let channel = |at: usize| u8::from_str_radix(&color[at..at + 2], 16).unwrap_or(0);
    let blend = |c: u8| -> u8 { (f64::from(c) * opacity + 255.0 * (1.0 - opacity)).round() as u8 };
    format!(
        "{:02x}{:02x}{:02x}",
        blend(channel(0)),
        blend(channel(2)),
        blend(channel(4))
    )
}

/// The colour LibreOffice paints automatic text in over an opaque
/// `background` (hex rrggbb): white when `Color::IsDark()`, else black
/// (`SwDrawTextInfo::ApplyAutoColor`, sw/source/core/txtnode/fntcache.cxx).
/// `IsDark()` (tools/source/generic/color.cxx, LibreOffice 25.2 to 26.8):
/// the WCAG relative luminance scaled to 255 and truncated, at most 87;
/// LibreOffice's default shape fill #729fcf instead uses the integer
/// luminance (76 R + 151 G + 29 B) >> 8, at most 62. Measured identical on
/// 953 backgrounds. The DOCX automatic colour uses another threshold, so the
/// reader writes the resolved colour.
pub(super) fn libreoffice_automatic_color(background: &str) -> &'static str {
    let channel = |at: usize| {
        background
            .get(at..at + 2)
            .and_then(|hex| u8::from_str_radix(hex, 16).ok())
            .unwrap_or(255)
    };
    let (r, g, b) = (channel(0), channel(2), channel(4));
    let dark = if (r, g, b) == (0x72, 0x9f, 0xcf) {
        (29 * u32::from(b) + 151 * u32::from(g) + 76 * u32::from(r)) >> 8 <= 62
    } else {
        let linear = |c: u8| {
            let v = f64::from(c) / 255.0;
            if v < 0.04045 {
                v / 12.92
            } else {
                ((v + 0.055) / 1.055).powf(2.4)
            }
        };
        let luminance = linear(r) * 0.2126 + linear(g) * 0.7152 + linear(b) * 0.0722;
        (luminance * 255.0) as u8 <= 87
    };
    if dark {
        "ffffff"
    } else {
        "000000"
    }
}
