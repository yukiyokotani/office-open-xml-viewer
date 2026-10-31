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
use crate::script::{is_weak_unmapped, table_22, ScriptClass};
use crate::style::{
    Cascade, Family, Group, Scope, Style, StyleRef, StyleSheet, LOEXT_GRAPHIC, PARAGRAPH, TEXT,
};
use crate::xml::{
    self, NS_DRAW, NS_FO, NS_LOEXT, NS_OFFICE, NS_OFFICEOOO, NS_STYLE, NS_SVG, NS_TEXT, NS_XML,
};
use docx_model::{
    BodyElement, BreakType, DocParagraph, DocRun, Document, DocumentSettings, LineSpacing,
    RunFontFacts, SectionProps, TextRun,
};
use std::cell::{Cell, RefCell};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::rc::Rc;

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
/// Implementation bound on the characters of one paragraph (after `text:s`
/// expansion). It bounds the transient token buffer of §6.1.2 normalization
/// (12 bytes per character) independently of the document-wide text budget.
const MAX_PARAGRAPH_CHARS: usize = 1_000_000;
/// Implementation bound on distinct span contexts in one paragraph.
const MAX_PARAGRAPH_CONTEXTS: usize = 65_536;
/// ODT-owned construction policy (a library policy, not a format limit):
/// the estimated native size of the retained model (text, runs, paragraphs),
/// charged before each allocation. The DOCX cursor additionally enforces its
/// retained JSON ceiling while serializing.
const MAX_CONSTRUCTION_BYTES: usize = 256 * 1024 * 1024;
/// Implementation bound on downstream layout work the admitted model causes,
/// document-wide. The DOCX line breaker sizes each line break by searching
/// the paragraph's runs for its index and then the nearest text run, so a
/// break costs about its run index plus the breaks it scans past.
const MAX_BREAK_LAYOUT_WORK: usize = 64 * 1024 * 1024;
/// Implementation bound on the §16.2 search order of one span: its own style
/// chain and those of the enclosing spans and paragraph.
const MAX_SEARCH_ORDER: usize = 256;
/// Implementation bound on cascade work, document-wide: every paragraph and
/// run resolution visits each style of its search order and each attribute
/// and element of the resolved property group (one unit each) and may process
/// each attribute's expanded name and value (one unit per byte), charged
/// before the resolution is built. Building a style's parent chain charges
/// one unit per identifier byte looked up.
const MAX_CASCADE_WORK: usize = 16 * 1024 * 1024;

/// Observed LibreOffice behaviour where ODF leaves the consumer free or
/// LibreOffice deviates from the text. Each is a user-approved LibreOffice
/// compatibility rule (2026-10-31, AGENTS.md) applied only within its
/// measured scope; outside it the reader rejects. Tests also run the reader
/// with a rule off to pin the spec-only behaviour.
#[derive(Clone, Copy, Debug)]
pub struct CompatRules {
    /// `fo:line-height` percentages scale the natural line height (DOCX
    /// "auto" with the same multiple). Admitted 100-300% on paragraphs whose
    /// text is one face and one size and not raised (glyph placement differs
    /// below 100%); other paragraphs and percentages reject.
    pub proportional_line_height: bool,
    /// Weak characters outside Table 22 (general categories Z*, P*, S*, Co
    /// and emoji-sequence components) take the preceding character's script
    /// class, or the following one's at paragraph start. Unmapped letters,
    /// numbers and other marks keep rejecting; digits after complex text
    /// (which LibreOffice sets in the complex font) are unreachable while
    /// complex text rejects.
    pub weak_script_class: bool,
    /// The nearest style that sets `fo:color` or
    /// `style:use-window-font-color="true"` decides the font colour
    /// (LibreOffice keeps one colour property). ODF inherits the two
    /// attributes separately (§16.2), and §20.395 makes an inherited
    /// automatic flag apply, so cascades where the two readings differ
    /// reject while this is off. A style setting both rejects.
    pub nearest_font_color: bool,
}

/// The rules in force: all three user-approved LibreOffice compatibility
/// rules (2026-10-31).
pub const LIBREOFFICE_RULES: CompatRules = CompatRules {
    proportional_line_height: true,
    weak_script_class: true,
    nearest_font_color: true,
};

pub fn read(package: &OdfPackage) -> Result<Document, String> {
    read_with(package, LIBREOFFICE_RULES)
}

fn read_with(package: &OdfPackage, rules: CompatRules) -> Result<Document, String> {
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
    let layout_flags = if package.contains("settings.xml") {
        check_settings(&package.read_text("settings.xml")?, version)?
    } else {
        LayoutFlags::default()
    };
    if package.contains("meta.xml") {
        let text = package.read_text("meta.xml")?;
        check_part_root(&xml::parse(&text)?, "document-meta", version)?;
    }
    // ODF 1.3 introduced style:contextual-spacing (§20.255); a declared 1.2
    // document does not define it.
    let contextual_spacing_defined = version != Some("1.2");
    // A background on the paragraph default style fills the whole page in
    // LibreOffice (measured), even under paragraphs that override it.
    if sheet
        .default_style(Family::Paragraph)
        .and_then(|style| style.group(PARAGRAPH))
        .and_then(|group| {
            group
                .attributes
                .get(&(NS_FO.to_string(), "background-color".to_string()))
        })
        .is_some_and(|value| value != "transparent")
    {
        return Err(unsupported("default-style-background"));
    }
    let page = PageStyle::read(styles_root)?;

    let body = content_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|b| b.children().find(|n| n.is_element()))
        .filter(|n| xml::is(*n, NS_OFFICE, "text"))
        .ok_or_else(|| unsupported("not-a-text-document"))?;

    let default_tab_stop = sheet
        .default_style(Family::Paragraph)
        .and_then(|style| Cascade::new(vec![style], PARAGRAPH).get(NS_STYLE, "tab-stop-distance"))
        .map(length_pt)
        .transpose()?;
    let mut reader = Reader {
        lists: list::Lists::default(),
        list_paragraph: None,
        odf_1_4: version == Some("1.4"),
        text_width: page.section.page_width - page.section.margin_left - page.section.margin_right,
        page_left_margin: page.section.margin_left,
        sheet: &sheet,
        fonts: &fonts,
        used_faces: RefCell::new(HashSet::new()),
        chains: HashMap::new(),
        text_chains: HashMap::new(),
        cascade_work: Cell::new(0),
        area: 0,
        next_area: 0,
        break_layout_work: 0,
        layout_flags,
        page_master: page.name.as_str(),
        paragraphs: Vec::new(),
        retained_text_bytes: 0,
        construction_bytes: 0,
        contextual_spacing_defined,
        rules,
    };
    reader.walk_blocks(body, 0)?;
    // The font table and its three metadata maps each copy a used family
    // name; charged before they are built.
    let font_table_bytes = reader
        .used_faces
        .borrow()
        .iter()
        .map(|face| 4 * (face.family.len() + 64))
        .sum::<usize>();
    reader.charge_construction(font_table_bytes)?;
    let used = used_font_table(&reader.used_faces.borrow())?;
    let body = reader.finish()?;

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
    /// Interned property strings (face families, languages) shared by every
    /// resolution that uses them.
    atoms: RefCell<HashSet<Rc<str>>>,
}

/// An interned string: equal atoms are the same allocation, so cached run
/// properties share them and compare and hash by identity in O(1).
#[derive(Clone, Debug)]
struct Atom(Rc<str>);

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

impl FontFaces {
    fn atom(&self, value: &str) -> Atom {
        let mut atoms = self.atoms.borrow_mut();
        if let Some(existing) = atoms.get(value) {
            return Atom(existing.clone());
        }
        let atom: Rc<str> = value.into();
        atoms.insert(atom.clone());
        Atom(atom)
    }
}

#[derive(Clone, PartialEq, Eq, Hash)]
struct Face {
    family: Atom,
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
            // Descriptors that select a face variant (§16.23: svg:font-style,
            // font-weight, font-variant, font-stretch, adornments, …) are not
            // mapped; only naming and font-table metadata are admitted.
            for attribute in face.attributes() {
                match (attribute.namespace(), attribute.name()) {
                    (
                        Some(NS_STYLE),
                        "name" | "font-family-generic" | "font-pitch" | "font-charset",
                    ) => {}
                    (Some(NS_SVG), "font-family") => {}
                    (Some(NS_STYLE), "font-adornments") if attribute.value().is_empty() => {}
                    _ => return Err(unsupported("font-face-descriptor")),
                }
            }
            let name = face
                .attribute((NS_STYLE, "name"))
                .ok_or_else(|| invalid("a font face has no name"))?;
            let family = face
                .attribute((NS_SVG, "font-family"))
                .ok_or_else(|| invalid("a font face has no family"))?;
            let value = Face {
                family: self.atom(&unquote_family(family)?),
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
fn used_font_table(used: &HashSet<Face>) -> Result<HashMap<String, Face>, String> {
    let mut table: HashMap<String, Face> = HashMap::new();
    for face in used {
        if let Some(existing) = table.insert(face.family.to_string(), face.clone()) {
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
        .filter_map(|face| field(face).map(|value| (face.family.to_string(), value.to_string())))
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

/// Applied styles may carry only the property groups the reader resolves;
/// conditional styles (`style:map`, §16.3) and any other child reject.
fn check_groups(chain: &[StyleRef<'_>], allowed: &[&str]) -> Result<(), String> {
    for style in chain {
        for group in style.style().groups.keys() {
            if !allowed.contains(&group.as_str()) {
                return Err(unsupported(if group == "map" {
                    "conditional-style"
                } else {
                    "style-group"
                }));
            }
        }
    }
    Ok(())
}

/// LibreOffice configuration items (`settings.xml`, not standardised by ODF)
/// that change text layout, pinned to the values LibreOffice uses for new ODF
/// documents — the conditions under which this reader's measurements hold.
/// A document setting any of them otherwise is rejected; items not listed are
/// view, print or editing settings.
const LAYOUT_SETTINGS: [(&str, &str); 13] = [
    ("ParaSpaceMax", "false"),
    ("ParaSpaceMaxAtPages", "false"),
    ("AddExternalLeading", "true"),
    ("UseFormerLineSpacing", "false"),
    ("PropLineSpacingShrinksFirstLine", "true"),
    ("UnxForceZeroExtLeading", "false"),
    ("UseOldPrinterMetrics", "false"),
    ("PrinterIndependentLayout", "high-resolution"),
    ("IgnoreHiddenCharsForLineCalculation", "true"),
    ("UseVariableWidthNBSP", "false"),
    ("ApplyTextAttrToEmptyLineAtEndOfParagraph", "false"),
    ("SmallCapsPercentage66", "false"),
    ("CharacterCompressionType", "0"),
];

/// LibreOffice configuration items that are admitted with either value:
/// measured (LibreOffice PDF with the item patched into its own
/// `settings.xml`) to change nothing this reader admits, or to change only
/// content the reader rejects. Each is re-examined when that content is
/// admitted.
///
/// - Tabs (`text:tab` rejects): TabsRelativeToIndent, TabOverMargin,
///   TabOverflow, TabOverSpacing, AlignTabStopPosition.
/// - A document grid (rejected unless its mode is none):
///   MsWordCompGridMetrics.
/// - The empty line after a paragraph-final line break (rejected):
///   ApplyParagraphMarkFormatToEmptyLineAtEndOfParagraph.
/// - Spaces larger than the adjacent text (rejected, `space-formatting`),
///   the only case where these were measured to differ:
///   IgnoreTabsAndBlanksForLineCalculation, MsWordCompTrailingBlanks.
/// - MsWordUlTrailSpace: no measured difference in underlined trailing
///   spaces at line and paragraph ends.
const INERT_SETTINGS: [&str; 10] = [
    "TabsRelativeToIndent",
    "TabOverMargin",
    "TabOverflow",
    "TabOverSpacing",
    "AlignTabStopPosition",
    "MsWordCompGridMetrics",
    "ApplyParagraphMarkFormatToEmptyLineAtEndOfParagraph",
    "IgnoreTabsAndBlanksForLineCalculation",
    "MsWordCompTrailingBlanks",
    "MsWordUlTrailSpace",
];

/// LibreOffice configuration items that change justification, admitted with
/// either value and applied to justified paragraphs (see `LayoutFlags`).
const JUSTIFICATION_SETTINGS: [&str; 2] = [
    "DoNotJustifyLinesWithManualBreak",
    "JustifyLinesWithShrinking",
];

/// LibreOffice configuration items that change space advances, admitted with
/// either value and applied by `emit_runs` (see `LayoutFlags`).
const SPACE_SETTINGS: [&str; 1] = ["BalanceSpacesAndIdeographicSpaces"];

/// Part 3 §3.1.3: each package part's root element, carrying the document's
/// one declared version (§19.388).
fn check_part_root(
    document: &roxmltree::Document<'_>,
    name: &str,
    version: Option<&str>,
) -> Result<(), String> {
    let root = document.root_element();
    if !xml::is(root, NS_OFFICE, name) {
        return Err(invalid("a package part has an unexpected root"));
    }
    if root.attribute((NS_OFFICE, "version")) != version {
        return Err(unsupported("odf-version"));
    }
    Ok(())
}

/// The settings.xml items that decide whether spaces size LibreOffice's
/// lines and carry decorations (absent: false, LibreOffice's default).
#[derive(Clone, Copy, Debug, Default)]
struct LayoutFlags {
    /// IgnoreTabsAndBlanksForLineCalculation: no space sizes a line.
    ignore: bool,
    /// MsWordCompTrailingBlanks: spaces that end a paragraph or precede a
    /// line break neither size the line nor carry decorations.
    blanks: bool,
    /// MsWordUlTrailSpace: with `blanks`, underline and strike-through still
    /// cover one such space.
    underline: bool,
    /// DoNotJustifyLinesWithManualBreak: a justified line ended by a manual
    /// break is left unjustified (LibreOffice otherwise stretches it).
    no_justify_breaks: bool,
    /// JustifyLinesWithShrinking: justification may also compress spaces.
    shrink: bool,
    /// BalanceSpacesAndIdeographicSpaces: a U+0020 at a text portion's edge
    /// or next to another U+0020 advances half its Asian font size.
    balance: bool,
    /// TabsRelativeToIndent, where the document sets it: list grid stops
    /// count from the paragraph indent when true, from the text area when
    /// false (measured); `None` leaves them unknown.
    tabs_relative: Option<bool>,
}

fn check_settings(text: &str, version: Option<&str>) -> Result<LayoutFlags, String> {
    let document = xml::parse(text)?;
    check_part_root(&document, "document-settings", version)?;
    let mut trailing = LayoutFlags::default();
    // LibreOffice reads its layout items from the top-level
    // "ooo:configuration-settings" set only; other sets are ignored.
    let mut sets = document
        .root_element()
        .children()
        .filter(|n| xml::is(*n, NS_OFFICE, "settings"))
        .flat_map(|settings| settings.children())
        .filter(|n| {
            xml::is(*n, xml::NS_CONFIG, "config-item-set")
                && n.attribute((xml::NS_CONFIG, "name")) == Some("ooo:configuration-settings")
        });
    let Some(set) = sets.next() else {
        return Ok(trailing);
    };
    if sets.next().is_some() {
        return Err(unsupported("compatibility-setting"));
    }
    let mut seen = HashSet::new();
    for item in set
        .children()
        .filter(|n| xml::is(*n, xml::NS_CONFIG, "config-item"))
    {
        let Some(name) = item.attribute((xml::NS_CONFIG, "name")) else {
            continue;
        };
        if !seen.insert(name) {
            return Err(unsupported("compatibility-setting"));
        }
        let value = item.text().unwrap_or("").trim();
        // §19.30 config:type; a boolean item must declare it.
        let boolean = || match (item.attribute((xml::NS_CONFIG, "type")), value) {
            (Some("boolean"), "true") => Ok(true),
            (Some("boolean"), "false") => Ok(false),
            _ => Err(unsupported("compatibility-setting")),
        };
        if let Some((_, expected)) = LAYOUT_SETTINGS.iter().find(|(key, _)| *key == name) {
            let matches = match *expected {
                "true" | "false" => boolean()? == (*expected == "true"),
                _ => value == *expected,
            };
            if !matches {
                return Err(unsupported("compatibility-setting"));
            }
        } else if INERT_SETTINGS.contains(&name)
            || JUSTIFICATION_SETTINGS.contains(&name)
            || SPACE_SETTINGS.contains(&name)
        {
            let value = boolean()?;
            match name {
                "IgnoreTabsAndBlanksForLineCalculation" => trailing.ignore = value,
                "MsWordCompTrailingBlanks" => trailing.blanks = value,
                "MsWordUlTrailSpace" => trailing.underline = value,
                "DoNotJustifyLinesWithManualBreak" => trailing.no_justify_breaks = value,
                "JustifyLinesWithShrinking" => trailing.shrink = value,
                "BalanceSpacesAndIdeographicSpaces" => trailing.balance = value,
                "TabsRelativeToIndent" => trailing.tabs_relative = Some(value),
                _ => {}
            }
        }
    }
    Ok(trailing)
}

/// Decorations over the spaces that end a paragraph or precede a line break
/// (measured, LibreOffice PDF against the DOCX layout). The DOCX layout paints
/// none over these hanging spaces (WORD_LINE_END_SPACE_HANGS; the reader sets
/// no ulTrailSpace). LibreOffice paints background, underline and
/// strike-through over them by default; with MsWordCompTrailingBlanks it
/// paints none, except underline and strike-through over one space with
/// MsWordUlTrailSpace as well. Decorated trailing spaces reject wherever the
/// two differ.
fn trailing_space_decoration(space: &RunProperties, settings: LayoutFlags) -> Result<(), String> {
    let differs = if settings.blanks {
        settings.underline && (space.underline.is_some() || space.strike.is_some())
    } else {
        decorated(space)
    };
    if differs {
        Err(unsupported("trailing-space-decoration"))
    } else {
        Ok(())
    }
}

fn decorated(space: &RunProperties) -> bool {
    space.background.is_some() || space.strike.is_some() || space.underline.is_some()
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
        // §19.845: text:number-lines defaults to true.
        if xml::is(node, NS_TEXT, "linenumbering-configuration")
            && node.attribute((NS_TEXT, "number-lines")) != Some("false")
        {
            return Err(unsupported("line-numbering"));
        }
        // §16.2 ends a text style's search at the text-family default style.
        // LibreOffice keeps text defaults in the paragraph default style and
        // writes no text-family default; one would change the cascade order
        // the reader implements, so it rejects.
        if xml::is(node, NS_STYLE, "default-style")
            && node.attribute((NS_STYLE, "family")) == Some("text")
            && node.children().any(|n| n.is_element())
        {
            return Err(unsupported("text-default-style"));
        }
        if xml::is(node, NS_TEXT, "outline-style") {
            for level in node.children().filter(|n| n.is_element()) {
                check_outline_level(level)?;
            }
        }
    }
    Ok(())
}

/// Headings are numbered and indented by the outline style (§16.36,
/// §16.37). A level is admitted only when it produces no label and its list
/// level properties (§17.19, §17.20) add no indent or spacing.
fn check_outline_level(level: roxmltree::Node<'_, '_>) -> Result<(), String> {
    if !xml::is(level, NS_TEXT, "outline-level-style") {
        return Err(unsupported("outline-style"));
    }
    for attribute in level.attributes() {
        let admitted = match (attribute.namespace(), attribute.name()) {
            (Some(NS_TEXT), "level" | "display-levels" | "start-value" | "style-name") => true,
            (Some(NS_STYLE), "num-letter-sync") => true,
            (Some(NS_STYLE), "num-format" | "num-prefix" | "num-suffix") => {
                if !attribute.value().is_empty() {
                    return Err(unsupported("heading-numbering"));
                }
                true
            }
            _ => false,
        };
        if !admitted {
            return Err(unsupported("outline-style"));
        }
    }
    let zero = |value: &str| length_pt(value).is_ok_and(|v| v == 0.0);
    for properties in level.children().filter(|n| n.is_element()) {
        if !xml::is(properties, NS_STYLE, "list-level-properties") {
            return Err(unsupported("outline-style"));
        }
        for attribute in properties.attributes() {
            let admitted = match (attribute.namespace(), attribute.name()) {
                (Some(NS_TEXT), "list-level-position-and-space-mode") => matches!(
                    attribute.value(),
                    "label-alignment" | "label-width-and-position"
                ),
                (Some(NS_TEXT), "space-before" | "min-label-width" | "min-label-distance") => {
                    zero(attribute.value())
                }
                _ => false,
            };
            if !admitted {
                return Err(unsupported("outline-style"));
            }
        }
        for alignment in properties.children().filter(|n| n.is_element()) {
            if !xml::is(alignment, NS_STYLE, "list-level-label-alignment")
                || alignment.children().any(|n| n.is_element())
            {
                return Err(unsupported("outline-style"));
            }
            for attribute in alignment.attributes() {
                let admitted = match (attribute.namespace(), attribute.name()) {
                    // What follows the (empty) label; a tab or space after
                    // an empty label is not measured.
                    (Some(NS_TEXT), "label-followed-by") => attribute.value() == "nothing",
                    (Some(NS_FO), "margin-left" | "text-indent") => zero(attribute.value()),
                    _ => false,
                };
                if !admitted {
                    return Err(unsupported("outline-style"));
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
        let mut size = (None, None);
        // §16.8: properties a page layout leaves unspecified come from the
        // default page layout.
        if let Some(default) = styles_root
            .children()
            .find(|n| xml::is(*n, NS_OFFICE, "styles"))
            .and_then(|s| {
                s.children()
                    .find(|n| xml::is(*n, NS_STYLE, "default-page-layout"))
            })
        {
            for child in default.children().filter(|n| n.is_element()) {
                match (child.tag_name().namespace(), child.tag_name().name()) {
                    (Some(NS_STYLE), "page-layout-properties") => {
                        read_page_properties(child, &mut section, &mut size)?
                    }
                    (Some(NS_STYLE), "header-style" | "footer-style") => {}
                    _ => return Err(unsupported("page-layout-child")),
                }
            }
        }
        let mut saw_properties = false;
        for child in layout.children().filter(|n| n.is_element()) {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_STYLE), "page-layout-properties") => {
                    read_page_properties(child, &mut section, &mut size)?;
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
        section.page_width = size
            .0
            .ok_or_else(|| invalid("a page layout has no width"))?;
        section.page_height = size
            .1
            .ok_or_else(|| invalid("a page layout has no height"))?;
        // Page dimensions are positive lengths (§20.215, §20.216) and the
        // margins (§20.204 et al.) must leave a text area.
        let s = &section;
        if !(s.page_width > 0.0 && s.page_height > 0.0)
            || [s.margin_top, s.margin_bottom, s.margin_left, s.margin_right]
                .iter()
                .any(|m| *m < 0.0)
            || s.margin_left + s.margin_right >= s.page_width
            || s.margin_top + s.margin_bottom >= s.page_height
        {
            return Err(invalid("a page layout has no text area"));
        }
        Ok(Self { name, section })
    }
}

fn read_page_properties(
    node: roxmltree::Node<'_, '_>,
    section: &mut SectionProps,
    size: &mut (Option<f64>, Option<f64>),
) -> Result<(), String> {
    let (width, height) = size;
    for attribute in node.attributes() {
        let value = attribute.value();
        match (attribute.namespace().unwrap_or(""), attribute.name()) {
            (NS_FO, "page-width") => *width = Some(length_pt(value)?),
            (NS_FO, "page-height") => *height = Some(length_pt(value)?),
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
            // Grid parameters are inert while the grid is off (§20.309).
            (NS_STYLE, name) if name.starts_with("layout-grid-") && name != "layout-grid-mode" => {}
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
    Ok(())
}

/// One collected paragraph before inter-paragraph spacing is resolved.
struct Collected {
    paragraph: DocParagraph,
    top: f64,
    bottom: f64,
    contextual: bool,
    /// The named (common) style compared by contextual spacing, by identity
    /// (common styles are unique per family and name).
    named_style: Option<*const Style>,
    /// The content area (body or `text:section`) holding the paragraph:
    /// §20.255 suppresses spacing only within one content area.
    area: u32,
    page_break_after: bool,
}

#[path = "odt_list.rs"]
mod list;

struct Reader<'a> {
    rules: CompatRules,
    /// List numbering state, and the list paragraph `paragraph` reads next.
    lists: list::Lists,
    list_paragraph: Option<list::ListParagraph<'a>>,
    /// The document declares ODF 1.4 (letter numbering beyond "z").
    odf_1_4: bool,
    /// The page's text area width and left margin (pt).
    text_width: f64,
    page_left_margin: f64,
    sheet: &'a StyleSheet,
    fonts: &'a FontFaces,
    used_faces: RefCell<HashSet<Face>>,
    /// Each applied style's parent chain, by style identity.
    chains: HashMap<*const Style, Rc<[StyleRef<'a>]>>,
    /// Each applied text style's checked parent chain, by style identity.
    text_chains: HashMap<*const Style, Rc<[StyleRef<'a>]>>,
    cascade_work: Cell<usize>,
    /// The current content area and the last one assigned.
    area: u32,
    next_area: u32,
    break_layout_work: usize,
    layout_flags: LayoutFlags,
    page_master: &'a str,
    paragraphs: Vec<Collected>,
    retained_text_bytes: usize,
    construction_bytes: usize,
    contextual_spacing_defined: bool,
}

/// One formatting context: a span's own style chain followed, through
/// `parent`, by the enclosing contexts (§16.2). Ancestors are shared by
/// index and each text style's chain is built once per document, so spans
/// copy no styles.
struct TextContext<'a> {
    styles: Rc<[StyleRef<'a>]>,
    parent: Option<u32>,
}

/// One paragraph's flattened inline content. Each token refers to its
/// formatting context by index into `contexts` (one entry per span).
struct Inline<'a> {
    contexts: Vec<TextContext<'a>>,
    /// Context index by (enclosing context, span style name): repeated spans
    /// share one context and so one property resolution.
    by_span: HashMap<(u32, String), u32>,
    tokens: Vec<Token>,
}

#[derive(Clone, Copy)]
enum Token {
    Char(char, u32),
    /// `text:s`: one space that is never collapsed.
    Space(u32),
    LineBreak(u32),
    /// No character, but where LibreOffice may end its text portion
    /// (BalanceSpacesAndIdeographicSpaces): `true` for a bookmark or
    /// reference mark (measured), `false` for an empty formatted span, whose
    /// formatting LibreOffice may keep as an empty attribute (not measured).
    Boundary(bool),
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
                (Some(NS_TEXT), "list") => self.list(node, None, depth)?,
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
                    let enclosing = self.area;
                    self.next_area += 1;
                    self.area = self.next_area;
                    self.walk_blocks(node, depth + 1)?;
                    self.area = enclosing;
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
        let list = self.list_paragraph.take();
        if list.is_some()
            && (node.attribute((NS_TEXT, "is-list-header")).is_some()
                || node.attribute((NS_TEXT, "restart-numbering")).is_some())
        {
            return Err(unsupported("list-paragraph-attribute"));
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
            Some(name) => {
                let style = self
                    .sheet
                    .find(Scope::Content, Family::Paragraph, name)
                    .ok_or_else(|| invalid("a paragraph names a missing style"))?;
                self.style_chain(style)?.to_vec()
            }
            None => Vec::new(),
        };
        let named_style = chain
            .iter()
            .find(|s| !s.style().automatic)
            .map(|s| s.style() as *const Style);
        if let Some(default) = self.sheet.default_style(Family::Paragraph) {
            chain.push(default);
        }
        self.charge_cascade(&chain, PARAGRAPH)?;
        let properties = Cascade::new(chain.clone(), PARAGRAPH);

        // §19.501: naming the page style on a later paragraph starts a page.
        let master = chain.iter().find_map(|s| {
            s.style()
                .attributes
                .get(&(NS_STYLE.to_string(), "master-page-name".to_string()))
                .map(String::as_str)
        });
        let mut page_break_before = false;
        if let Some(master) = master.filter(|name| !name.is_empty()) {
            self.charge_work(master.len())?;
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

        check_groups(
            &chain,
            &["paragraph-properties", "text-properties", LOEXT_GRAPHIC],
        )?;
        self.charge_cascade(&chain, LOEXT_GRAPHIC)?;
        check_paragraph_fill(&properties, &Cascade::new(chain.clone(), LOEXT_GRAPHIC))?;
        let mut collected = paragraph_properties(&properties, self.rules)?;
        if collected.contextual && !self.contextual_spacing_defined {
            return Err(unsupported("contextual-spacing-version"));
        }
        collected.named_style = named_style;
        collected.area = self.area;
        collected.paragraph.page_break_before |= page_break_before;
        collected.paragraph.outline_level = outline_level.map(|level| level.saturating_sub(1));
        if let Some(list) = &list {
            self.list_geometry(list, &chain, &properties, &mut collected.paragraph)?;
        }

        // The paragraph's own text properties describe its mark (empty
        // paragraph height) and the text outside spans.
        let contexts = vec![TextContext {
            styles: chain.into(),
            parent: None,
        }];
        let mark = self.run_properties(&contexts, 0, ScriptClass::Latin)?;
        // The paragraph's retained size, charged before its strings (five
        // copies of the mark family) are allocated.
        self.charge_construction(
            std::mem::size_of::<Collected>()
                + std::mem::size_of::<BodyElement>()
                + 5 * mark.family.len(),
        )?;
        collected.paragraph.paragraph_mark_font_facts = Some(mark.font_facts());
        collected.paragraph.default_font_size = Some(mark.size);
        collected.paragraph.default_font_family = Some(mark.family.to_string());

        let mut inline = Inline {
            contexts,
            by_span: HashMap::new(),
            tokens: Vec::new(),
        };
        self.inline(node, 0, &mut inline, 0)?;
        if collected.paragraph.alignment == "justify" {
            // Justified lines (DOCX `both`, WORD_LINE_END_SPACE_HANGS) match
            // LibreOffice's breaks and flush edges (measured; LibreOffice's
            // whole-unit distribution leaves its lines up to 0.1 pt per space
            // short). JustifyLinesWithShrinking also compresses spaces, and
            // LibreOffice stretches a line ended by a manual break unless
            // DoNotJustifyLinesWithManualBreak; the DOCX layout does neither.
            if self.layout_flags.shrink {
                return Err(unsupported("justify-shrinking"));
            }
            if !self.layout_flags.no_justify_breaks
                && inline
                    .tokens
                    .iter()
                    .any(|token| matches!(token, Token::LineBreak(_)))
            {
                return Err(unsupported("justify-with-line-break"));
            }
        }
        let (tokens, removed) = collapse(std::mem::take(&mut inline.tokens));
        let justified = collected.paragraph.alignment == "justify";
        self.emit_runs(
            &tokens,
            removed,
            &inline.contexts,
            justified,
            &mut collected.paragraph,
        )?;
        if let Some(label) = list.and_then(|list| list.label) {
            self.list_label(label, &mut inline.contexts, &mut collected.paragraph)?;
        }
        if let Some(shading) = collected.paragraph.shading.clone() {
            // §20.395 automatic colour (also an unset colour) over the
            // paragraph background, resolved as LibreOffice paints it (runs
            // with their own shading were resolved against it).
            let resolved = libreoffice_automatic_color(&shading)?;
            for run in &mut collected.paragraph.runs {
                if let DocRun::Text(text) = run {
                    if text.color.is_none() && text.background.is_none() {
                        text.color = Some(resolved.to_string());
                        text.color_auto = false;
                    }
                }
            }
        }
        if collected.paragraph.line_spacing.is_some() {
            // Proportional spacing scales the natural line of uniform text
            // alike; with mixed faces or sizes LibreOffice does not multiply
            // the extra height (measured), so such paragraphs reject.
            let paragraph = &collected.paragraph;
            let uniform = paragraph.runs.iter().all(|run| match run {
                DocRun::Text(text) => {
                    text.font_family == paragraph.default_font_family
                        && Some(text.font_size) == paragraph.default_font_size
                }
                _ => true,
            });
            if !uniform {
                return Err(unsupported("line-height-proportional-mixed"));
            }
        }
        self.paragraphs.push(collected);
        Ok(())
    }

    /// Inline content (§6.1): character data, spans and the white-space
    /// elements, flattened into tokens for §6.1.2 collapsing.
    fn inline(
        &mut self,
        node: roxmltree::Node<'_, '_>,
        context: u32,
        inline: &mut Inline<'a>,
        depth: usize,
    ) -> Result<(), String> {
        if depth > MAX_INLINE_DEPTH {
            return Err(unsupported("inline-depth"));
        }
        for child in node.children() {
            if child.is_text() {
                let text = child.text().unwrap_or("");
                self.charge_text(text.len(), text.chars().count(), inline)?;
                inline
                    .tokens
                    .extend(text.chars().map(|c| Token::Char(c, context)));
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
                        Some(name) => match inline.by_span.get(&(context, name.to_string())) {
                            Some(index) => *index,
                            None => {
                                let style = self
                                    .sheet
                                    .find(Scope::Content, Family::Text, name)
                                    .ok_or_else(|| invalid("a span names a missing style"))?;
                                if inline.contexts.len() >= MAX_PARAGRAPH_CONTEXTS {
                                    return Err(unsupported("span-count"));
                                }
                                let styles = self.text_chain(style)?;
                                inline.contexts.push(TextContext {
                                    styles,
                                    parent: Some(context),
                                });
                                let index = (inline.contexts.len() - 1) as u32;
                                inline.by_span.insert((context, name.to_string()), index);
                                index
                            }
                        },
                        None => context,
                    };
                    let before = inline.tokens.len();
                    self.inline(child, nested, inline, depth + 1)?;
                    if self.layout_flags.balance
                        && inline.tokens.len() == before
                        && nested != context
                        && span_formats(inline.contexts[nested as usize].styles[0].style())
                    {
                        self.charge_text(1, 1, inline)?;
                        inline.tokens.push(Token::Boundary(false));
                    }
                }
                (Some(NS_TEXT), "s") => {
                    let count = child
                        .attribute((NS_TEXT, "c"))
                        .map(|v| v.parse::<usize>().map_err(|_| invalid("bad text:s count")))
                        .transpose()?
                        .unwrap_or(1);
                    if count == 0 {
                        return Err(invalid("text:s count is out of range"));
                    }
                    // Expanded spaces are retained text like character data.
                    self.charge_text(count, count, inline)?;
                    inline
                        .tokens
                        .extend((0..count).map(|_| Token::Space(context)));
                }
                // Tab geometry (positions from the indent, the default
                // interval, synthetic tab metrics, decorations across the gap
                // and tab-only wrapped lines) differs between LibreOffice and
                // the DOCX layout in several measured ways; tabs reject until
                // they are mapped as a whole.
                (Some(NS_TEXT), "tab") => return Err(unsupported("tab")),
                (Some(NS_TEXT), "line-break") => {
                    self.charge_text(1, 1, inline)?;
                    inline.tokens.push(Token::LineBreak(context));
                }
                (Some(NS_TEXT), "soft-page-break") => {}
                (
                    Some(NS_TEXT),
                    "bookmark"
                    | "bookmark-start"
                    | "bookmark-end"
                    | "reference-mark"
                    | "reference-mark-start"
                    | "reference-mark-end",
                ) => {
                    self.charge_text(1, 1, inline)?;
                    inline.tokens.push(Token::Boundary(true));
                }
                _ => return Err(unsupported("inline-element")),
            }
        }
        Ok(())
    }

    fn text_chain(&mut self, style: StyleRef<'a>) -> Result<Rc<[StyleRef<'a>]>, String> {
        let key: *const Style = style.style();
        if let Some(chain) = self.text_chains.get(&key) {
            return Ok(chain.clone());
        }
        let chain = self.style_chain(style)?;
        check_groups(&chain, &["text-properties"])?;
        self.text_chains.insert(key, chain.clone());
        Ok(chain)
    }

    /// A style's parent chain, built once per style and charged to the
    /// cascade work budget.
    fn style_chain(&mut self, style: StyleRef<'a>) -> Result<Rc<[StyleRef<'a>]>, String> {
        let key: *const Style = style.style();
        if let Some(chain) = self.chains.get(&key) {
            return Ok(chain.clone());
        }
        let chain: Rc<[StyleRef<'a>]> = self
            .sheet
            .chain(style, |work| self.charge_work(work))?
            .into();
        self.chains.insert(key, chain.clone());
        Ok(chain)
    }

    /// Charge retained text before the tokens are allocated: the
    /// document-wide byte budget and the per-paragraph character bound.
    fn charge_text(
        &mut self,
        bytes: usize,
        chars: usize,
        inline: &Inline<'a>,
    ) -> Result<(), String> {
        self.retained_text_bytes = self
            .retained_text_bytes
            .checked_add(bytes)
            .filter(|total| *total <= MAX_RETAINED_TEXT_BYTES)
            .ok_or_else(|| "OpenDocument text exceeds the retained text limit".to_string())?;
        self.charge_construction(bytes)?;
        if inline.tokens.len().saturating_add(chars) > MAX_PARAGRAPH_CHARS {
            return Err("OpenDocument paragraph exceeds the character limit".to_string());
        }
        Ok(())
    }

    fn charge_construction(&mut self, bytes: usize) -> Result<(), String> {
        self.construction_bytes = self
            .construction_bytes
            .checked_add(bytes)
            .filter(|total| *total <= MAX_CONSTRUCTION_BYTES)
            .ok_or_else(|| "OpenDocument text exceeds the construction limit".to_string())?;
        Ok(())
    }

    /// BalanceSpacesAndIdeographicSpaces (LibreOffice `sw::Justify::
    /// BalanceCjkSpaces`, measured): a U+0020 that is the first or last
    /// character of a text portion, or next to another U+0020, advances half
    /// its Asian font size, unchanged by horizontal scaling and with letter spacing
    /// added. Portions end at the paragraph's ends, line breaks, soft hyphens,
    /// bookmarks and reference marks, and where the span formatting changes
    /// (`Self::seam`). The break search uses the same advances, except in
    /// justified paragraphs, where LibreOffice measured some of them unbalanced
    /// and then compressed the line; a justified paragraph rejects unless
    /// every balanced space hangs at its line's end. Returns each token's
    /// advance in points (`None` for other characters).
    fn balanced_spaces(
        &mut self,
        tokens: &[Token],
        contexts: &[TextContext<'a>],
        justified: bool,
    ) -> Result<Vec<Option<f64>>, String> {
        let space =
            |token: Option<&Token>| matches!(token, Some(Token::Char(' ', _) | Token::Space(_)));
        let context = |token: &Token| match *token {
            Token::Char(_, context) | Token::Space(context) => Some(context),
            Token::LineBreak(_) | Token::Boundary(_) => None,
        };
        // Whether each token is followed only by U+0020 and portion marks up
        // to the next line break or the paragraph end (one reverse pass).
        let mut hangs = vec![false; tokens.len()];
        let mut rest_hangs = true;
        for (index, token) in tokens.iter().enumerate().rev() {
            hangs[index] = rest_hangs;
            rest_hangs = match token {
                Token::LineBreak(_) => true,
                Token::Boundary(_) => rest_hangs,
                token => rest_hangs && space(Some(token)),
            };
        }
        let mut seams = SeamCache::default();
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
            let mut unknown = false;
            for neighbour in [previous, next] {
                if balanced {
                    break;
                }
                match neighbour {
                    None | Some(Token::LineBreak(_) | Token::Boundary(true)) => balanced = true,
                    Some(Token::Boundary(false)) => unknown = true,
                    // Characters LibreOffice sets in portions of their own
                    // (measured: soft hyphen, no-break space, non-breaking
                    // hyphen, zero-width space, word joiner) and ones whose
                    // effect was not reproduced (byte order mark, direction
                    // controls).
                    Some(Token::Char(c, _)) if portion_character(*c) => balanced = true,
                    Some(Token::Char(c, _)) if uncertain_portion_character(*c) => unknown = true,
                    Some(other) => {
                        let other = context(other).expect("a character has a context");
                        match self.seam(&mut seams, contexts, own, other)? {
                            Seam::Same => {}
                            Seam::Boundary => balanced = true,
                            Seam::Unknown => unknown = true,
                        }
                    }
                }
            }
            if !balanced {
                if unknown {
                    return Err(unsupported("balanced-space-portion"));
                }
                continue;
            }
            if justified && !hangs[index] {
                return Err(unsupported("justify-balanced-spaces"));
            }
            let advance = match advances.get(&own) {
                Some(advance) => *advance,
                None => {
                    let advance = self.balanced_space_advance(contexts, own)?;
                    advances.insert(own, advance);
                    advance
                }
            };
            out[index] = Some(advance);
        }
        Ok(out)
    }

    /// Whether LibreOffice starts a text portion between two span contexts.
    /// LibreOffice compares the imported formatting items after normalising
    /// values, removing span properties that repeat the paragraph's direct
    /// (automatic style) formatting and applying nested character styles
    /// over enclosing automatic properties; that comparison is not
    /// reproduced. Instead: identical span chains are one portion; a
    /// different character style (by display name) or a clearly different
    /// resolved value (formatting the reader renders, or the Asian size) is
    /// a boundary (measured); anything else, such as a different language,
    /// is unknown.
    fn seam(
        &self,
        cache: &mut SeamCache<'a>,
        contexts: &[TextContext<'a>],
        a: u32,
        b: u32,
    ) -> Result<Seam, String> {
        if a == b {
            return Ok(Seam::Same);
        }
        let pair = (a.min(b), a.max(b));
        if let Some(seam) = cache.seams.get(&pair) {
            return Ok(*seam);
        }
        for context in [a, b] {
            if cache.chains.contains_key(&context) {
                continue;
            }
            let chain = span_chain(contexts, context);
            let work = chain.iter().fold(1usize, |work, (name, properties)| {
                properties.iter().fold(
                    work.saturating_add(name.map_or(0, str::len)),
                    |work, (key, value)| work.saturating_add(key.1.len() + value.len() + 1),
                )
            });
            self.charge_work(work)?;
            let properties = self.run_properties(contexts, context, ScriptClass::Latin)?;
            let asian = self.balanced_space_advance(contexts, context)? * 2.0;
            cache.chains.insert(context, chain);
            cache.resolved.insert(context, (properties, asian));
        }
        let seam = if cache.chains[&a] == cache.chains[&b] {
            Seam::Same
        } else {
            let names = |chain: &SpanChain<'a>| {
                let mut names: Vec<&str> = chain.iter().filter_map(|(name, _)| *name).collect();
                names.sort_unstable();
                names
            };
            let ((pa, aa), (pb, ab)) = (&cache.resolved[&a], &cache.resolved[&b]);
            // Languages are left out: LibreOffice maps locale aliases (zh and
            // cmn, a country without a language) together, so a language
            // difference alone stays unknown.
            if names(&cache.chains[&a]) != names(&cache.chains[&b]) || (aa - ab).abs() >= 0.1 {
                Seam::Boundary
            } else {
                // Sizes LibreOffice may round together, languages, and family
                // names differing only in case stay unknown.
                let mut near = pb.clone();
                if (pa.size - pb.size).abs() < 0.1 {
                    near.size = pa.size;
                }
                near.lang_east_asia = pa.lang_east_asia.clone();
                near.lang_bidi = pa.lang_bidi.clone();
                if pa.family.eq_ignore_ascii_case(&near.family) {
                    near.family = pa.family.clone();
                }
                if *pa != near {
                    Seam::Boundary
                } else {
                    Seam::Unknown
                }
            }
        };
        cache.seams.insert(pair, seam);
        Ok(seam)
    }

    /// Half the Asian font size (LibreOffice's "CJK font height × 0.5": the
    /// nominal size, without measuring a glyph).
    fn balanced_space_advance(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<f64, String> {
        let styles = self.search_order(contexts, context)?;
        self.charge_cascade(&styles, TEXT)?;
        let cascade = Cascade::new(styles, TEXT);
        let size = match cascade.get(NS_STYLE, "font-size-asian") {
            Some(value) if value.ends_with('%') => return Err(unsupported("relative-font-size")),
            Some(value) => match length_pt(value)? {
                size if (1.0..=1000.0).contains(&size) => size,
                _ => return Err(unsupported("font-size-range")),
            },
            None => return Err(unsupported("font-size-unresolved")),
        };
        // LibreOffice keeps font heights in twips (12.024 pt balances as
        // 12 pt, verified); sizes between twips are not reproduced.
        let twips = size * 20.0;
        if (twips - twips.round()).abs() > 1e-6 {
            return Err(unsupported("balanced-space-size"));
        }
        Ok(twips.round() / 40.0)
    }

    /// Script classes (§20.358 Table 22) and run emission.
    fn emit_runs(
        &mut self,
        tokens: &[Token],
        removed_space: Option<u32>,
        contexts: &[TextContext<'a>],
        justified: bool,
        paragraph: &mut DocParagraph,
    ) -> Result<(), String> {
        let balanced = if self.layout_flags.balance {
            self.balanced_spaces(tokens, contexts, justified)?
        } else {
            Vec::new()
        };
        // Properties are resolved once per (context, class); a run continues
        // while they are equal, so merging is linear in the text.
        let mut cache: HashMap<(u32, ScriptClass), RunProperties> = HashMap::new();
        let mut uniform: HashMap<u32, ScriptClass> = HashMap::new();
        let mut current: Option<(u32, ScriptClass)> = None;
        // Breaks since the last text run: each scans back past them.
        let mut breaks_since_text = 0usize;
        // Whether the current line holds a glyph with font metrics: the DOCX
        // layout gives lines of only white space synthetic metrics, unlike
        // LibreOffice.
        let mut line_has_glyph = false;
        let mut any_break = false;
        let weak_classes = if self.rules.weak_script_class {
            weak_classes(tokens)
        } else {
            Vec::new()
        };
        // The last text run, independent of run merging (`current`, which a
        // soft hyphen's own run interrupts).
        let mut last_text: Option<(u32, ScriptClass)> = None;
        // Soft hyphens before any text, checked against the first text run.
        let mut pending_hyphens: Vec<(u32, ScriptClass)> = Vec::new();
        // The DOCX layout sizes lines by every space, as LibreOffice does by
        // default (measured mid-line, at a wrap, before a break and at the
        // paragraph end). Where LibreOffice's settings make a space not size
        // its line (every space with IgnoreTabsAndBlanksForLineCalculation,
        // a trailing one with MsWordCompTrailingBlanks) the space must add
        // nothing to the DOCX line either: it must be in the face (family,
        // weight, posture) and raise/lower of the nearest visible character
        // on its line, preceding first, and no larger. Any space may end a
        // wrapped line, so the following character only stands in at a line
        // start. A manual break starts a new comparison.
        let mut last_glyph: Option<(u32, ScriptClass)> = None;
        let mut pending_spaces: Vec<(u32, ScriptClass)> = Vec::new();
        // White space since the last visible character, which ends the
        // paragraph or precedes a break if nothing visible follows.
        let mut trailing: Vec<(u32, ScriptClass, char)> = Vec::new();
        for (index, &token) in tokens.iter().enumerate() {
            let (c, context) = match token {
                Token::LineBreak(context) => {
                    // LibreOffice sizes a break's line portion by the break's
                    // own font (measured); the DOCX layout by the nearest
                    // preceding text run. They agree only when those fonts
                    // have the same line metrics.
                    let key = (context, ScriptClass::Latin);
                    if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(key) {
                        entry.insert(self.run_properties(contexts, context, key.1)?);
                    }
                    // A line holding only a break or white space (leading,
                    // consecutive or trailing breaks) gets synthetic metrics
                    // in the DOCX layout, unlike LibreOffice's full line
                    // (measured), so each line a break ends or starts needs a
                    // visible character.
                    if !line_has_glyph || !pending_spaces.is_empty() {
                        return Err(unsupported("line-break-empty-line"));
                    }
                    self.trailing_spaces(&cache, &std::mem::take(&mut trailing), last_glyph)?;
                    last_glyph = None;
                    line_has_glyph = false;
                    any_break = true;
                    match last_text {
                        Some(text) => {
                            same_line_metrics(&cache[&key], &cache[&text], "line-break-formatting")?
                        }
                        None => return Err(unsupported("line-break-empty-line")),
                    }
                    self.break_layout_work = self
                        .break_layout_work
                        .saturating_add(paragraph.runs.len() + breaks_since_text + 1);
                    if self.break_layout_work > MAX_BREAK_LAYOUT_WORK {
                        return Err("OpenDocument text exceeds the line break layout limit".into());
                    }
                    breaks_since_text += 1;
                    self.charge_construction(std::mem::size_of::<DocRun>())?;
                    paragraph.runs.push(DocRun::Break {
                        break_type: BreakType::Line,
                    });
                    continue;
                }
                Token::Boundary(_) => continue,
                Token::Char(c, context) => (c, context),
                Token::Space(context) => (' ', context),
            };
            // A character Table 22 does not map has an implementation-dependent
            // class (§20.358). LibreOffice was observed to use the preceding
            // character's class; until that rule is approved, such a character
            // is admitted only where every class resolves to the same face.
            let class = match table_22(c).or_else(|| weak_classes.get(index).copied().flatten()) {
                Some(class) => class,
                None => match uniform.get(&context) {
                    Some(class) => *class,
                    None => {
                        let class = self.uniform_class(contexts, context)?;
                        uniform.insert(context, class);
                        class
                    }
                },
            };
            if class != ScriptClass::Latin {
                // Asian and complex text (line breaking, punctuation, autospace
                // and shaping rules) is not yet measured against LibreOffice.
                return Err(unsupported("script-class"));
            }
            let key = (context, class);
            if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(key) {
                entry.insert(self.run_properties(contexts, context, class)?);
            }
            breaks_since_text = 0;
            if c == '\u{ad}' {
                // §6.1.6 soft hyphen: a break opportunity that shows a hyphen
                // in its own formatting only when the break is taken
                // (measured), as the DOCX optional hyphen run does. Its
                // formatting does not size LibreOffice's line but does size
                // the DOCX one, so it must share the adjacent text's metrics.
                match last_text {
                    Some(text) => {
                        same_line_metrics(&cache[&key], &cache[&text], "soft-hyphen-formatting")?
                    }
                    None => pending_hyphens.push(key),
                }
                self.charge_construction(cache[&key].retained_estimate())?;
                let mut run = cache[&key].text_run();
                run.optional_hyphen = true;
                paragraph.runs.push(DocRun::Text(Box::new(run)));
                current = None;
                continue;
            }
            for hyphen in pending_hyphens.drain(..) {
                same_line_metrics(&cache[&hyphen], &cache[&key], "soft-hyphen-formatting")?;
            }
            if c.is_whitespace() {
                if self.layout_flags.ignore {
                    match last_glyph {
                        Some(glyph) => space_fits_line(&cache[&key], &cache[&glyph])?,
                        None => pending_spaces.push(key),
                    }
                }
                trailing.push((key.0, key.1, c));
            } else {
                trailing.clear();
                for space in pending_spaces.drain(..) {
                    space_fits_line(&cache[&space], &cache[&key])?;
                }
                last_glyph = Some(key);
            }
            line_has_glyph |= !c.is_whitespace();
            // A balanced space keeps its own run, which carries its advance.
            let advance = balanced.get(index).copied().flatten();
            let continues = match (current, paragraph.runs.last_mut()) {
                (Some(last_key), Some(DocRun::Text(last)))
                    if cache[&last_key] == cache[&key] && last.space_advance_pt == advance =>
                {
                    last.text.push(c);
                    true
                }
                _ => false,
            };
            if !continues {
                self.charge_construction(cache[&key].retained_estimate())?;
                let mut run = cache[&key].text_run();
                run.text.push(c);
                run.space_advance_pt = advance;
                paragraph.runs.push(DocRun::Text(Box::new(run)));
            }
            current = Some(key);
            last_text = Some(key);
        }
        if any_break && !line_has_glyph {
            return Err(unsupported("line-break-empty-line"));
        }
        if !pending_hyphens.is_empty() {
            return Err(unsupported("soft-hyphen-formatting"));
        }
        if !pending_spaces.is_empty() {
            return Err(unsupported("space-formatting"));
        }
        self.trailing_spaces(&cache, &trailing, last_glyph)?;
        // §6.1.2 step 5 removes a trailing collapsed space; LibreOffice keeps
        // it (measured: it sizes the line and carries decorations unless
        // its settings say otherwise), so it must be invisible there too.
        if let Some(context) = removed_space {
            let key = (context, ScriptClass::Latin);
            if let std::collections::hash_map::Entry::Vacant(entry) = cache.entry(key) {
                entry.insert(self.run_properties(contexts, context, key.1)?);
            }
            let space = &cache[&key];
            let settings = self.layout_flags;
            if !settings.ignore && !settings.blanks {
                match last_glyph {
                    Some(glyph) => space_fits_line(space, &cache[&glyph])?,
                    None => return Err(unsupported("space-formatting")),
                }
            }
            // A U+0020 after every other trailing character: decorated as
            // a retained trailing space would be.
            trailing_space_decoration(space, settings)?;
        }
        Ok(())
    }

    /// The white space that ends a paragraph or precedes a line break, after
    /// `glyph` (the last visible character on its line). With
    /// MsWordCompTrailingBlanks LibreOffice ignores it for the line height,
    /// so the DOCX layout must too (see `emit_runs`); the IgnoreTabsAndBlanks
    /// case was already checked as an ordinary space. Only the U+0020 suffix
    /// hangs in the DOCX layout; LibreOffice, like the DOCX layout, decorates
    /// a trailing no-break or other Unicode space under every setting
    /// (measured: U+00A0, U+2002, U+2009), so only that suffix meets the
    /// decoration gate.
    fn trailing_spaces(
        &self,
        cache: &HashMap<(u32, ScriptClass), RunProperties>,
        trailing: &[(u32, ScriptClass, char)],
        glyph: Option<(u32, ScriptClass)>,
    ) -> Result<(), String> {
        if trailing.is_empty() {
            return Ok(());
        }
        let settings = self.layout_flags;
        // A line of only white space gets synthetic metrics in the DOCX
        // layout (see the line-break gate), whatever the settings.
        let glyph = &cache[&glyph.ok_or_else(|| unsupported("space-formatting"))?];
        let hanging = trailing.len()
            - trailing
                .iter()
                .rev()
                .take_while(|(_, _, c)| *c == ' ')
                .count();
        for (index, (context, class, _)) in trailing.iter().enumerate() {
            let space = &cache[&(*context, *class)];
            if settings.blanks && !settings.ignore {
                space_fits_line(space, glyph)?;
            }
            if index >= hanging {
                trailing_space_decoration(space, settings)?;
            }
        }
        Ok(())
    }

    fn uniform_class(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<ScriptClass, String> {
        let latin = self.run_properties(contexts, context, ScriptClass::Latin)?;
        let asian = self.run_properties(contexts, context, ScriptClass::Asian)?;
        let complex = self.run_properties(contexts, context, ScriptClass::Complex)?;
        if latin == asian && latin == complex {
            Ok(ScriptClass::Latin)
        } else {
            Err(unsupported("script-unmapped-character"))
        }
    }

    fn run_properties(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
        class: ScriptClass,
    ) -> Result<RunProperties, String> {
        let styles = self.search_order(contexts, context)?;
        self.charge_cascade(&styles, TEXT)?;
        let cascade = Cascade::new(styles, TEXT);
        let (properties, face) = RunProperties::resolve(&cascade, class, self.fonts, self.rules)?;
        self.used_faces.borrow_mut().insert(face);
        Ok(properties)
    }

    fn charge_cascade(&self, styles: &[StyleRef<'a>], group: Group) -> Result<(), String> {
        let work = styles.iter().fold(0usize, |work, style| {
            work.saturating_add(1)
                .saturating_add(style.group(group).map_or(0, |p| {
                    p.element_nodes
                        + p.attributes
                            .iter()
                            .map(|((namespace, name), value)| {
                                1 + namespace.len() + name.len() + value.len()
                            })
                            .sum::<usize>()
                }))
        });
        self.charge_work(work)
    }

    fn charge_work(&self, work: usize) -> Result<(), String> {
        let total = self
            .cascade_work
            .get()
            .checked_add(work)
            .filter(|total| *total <= MAX_CASCADE_WORK)
            .ok_or_else(|| "OpenDocument text exceeds the style resolution limit".to_string())?;
        self.cascade_work.set(total);
        Ok(())
    }

    /// The §16.2 search order of one context, nearest style first (at most
    /// `MAX_SEARCH_ORDER` references).
    fn search_order(
        &self,
        contexts: &[TextContext<'a>],
        context: u32,
    ) -> Result<Vec<StyleRef<'a>>, String> {
        let mut length = 0usize;
        let mut current = Some(context);
        while let Some(index) = current {
            let entry = &contexts[index as usize];
            length += entry.styles.len();
            current = entry.parent;
        }
        if length > MAX_SEARCH_ORDER {
            return Err("OpenDocument span nesting exceeds the style search order limit".into());
        }
        let mut styles = Vec::with_capacity(length);
        let mut current = Some(context);
        while let Some(index) = current {
            let entry = &contexts[index as usize];
            styles.extend(entry.styles.iter().copied());
            current = entry.parent;
        }
        Ok(styles)
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
        let mut previous: Option<Collected> = None;
        // The previous paragraph's background (its model has moved into the
        // output).
        let mut previous_shading: Option<String> = None;
        let mut break_pending = false;
        for mut collected in self.paragraphs {
            let mut paragraph = std::mem::take(&mut collected.paragraph);
            paragraph.page_break_before |= break_pending;
            paragraph.space_before = match previous {
                // Measured: after a hard page break (break-before, break-after
                // or a page style named on the paragraph) LibreOffice applies
                // the new paragraph's top margin alone; the previous bottom
                // margin is not carried. At an automatic page break both
                // vanish, as the DOCX layout drops space before at the top of
                // a page it reached by pagination.
                None => collected.top,
                Some(_) if paragraph.page_break_before => collected.top,
                Some(ref prev) => {
                    let same = prev.area == collected.area
                        && prev.named_style.is_some()
                        && prev.named_style == collected.named_style;
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
            if let Some(prev) = previous.as_ref() {
                // LibreOffice joins the backgrounds of consecutive paragraphs
                // of the same colour in one content area across the spacing
                // between them (measured: a 24 pt gap filled). Different
                // colours, or a section boundary, leave the gap unfilled, as
                // the DOCX layout does for every gap.
                if previous_shading.is_some()
                    && previous_shading == paragraph.shading
                    && prev.area == collected.area
                    && !paragraph.page_break_before
                    && paragraph.space_before > 0.0
                {
                    return Err(unsupported("paragraph-fill-gap"));
                }
            }
            previous_shading = paragraph.shading.clone();
            paragraph.space_after = 0.0;
            paragraph.contextual_spacing = false;
            break_pending = collected.page_break_after;
            out.push(BodyElement::Paragraph(Box::new(paragraph)));
            previous = Some(collected);
        }
        Ok(out)
    }
}

/// LibreOffice's class for weak characters outside Table 22 (measured;
/// `CompatRules::weak_script_class`): the class of the preceding character,
/// or at paragraph start of the following one. `None` where no rule applies
/// (other unmapped characters, or a paragraph without a mapped character).
fn weak_classes(tokens: &[Token]) -> Vec<Option<ScriptClass>> {
    let character = |token: &Token| match *token {
        Token::Char(c, _) => Some(c),
        Token::Space(_) => Some(' '),
        Token::LineBreak(_) | Token::Boundary(_) => None,
    };
    let mut classes: Vec<Option<ScriptClass>> = Vec::with_capacity(tokens.len());
    let mut previous: Option<ScriptClass> = None;
    for token in tokens {
        let class = match character(token) {
            Some(c) => match table_22(c) {
                Some(class) => Some(class),
                None if is_weak_unmapped(c) => previous,
                None => None,
            },
            None => None,
        };
        if class.is_some() {
            previous = class;
        }
        classes.push(class);
    }
    // Leading weak characters take the first following class.
    if let Some(first) = classes.iter().position(Option::is_some) {
        let following = classes[first];
        for (index, token) in tokens[..first].iter().enumerate() {
            if character(token).is_some_and(is_weak_unmapped) {
                classes[index] = following;
            }
        }
    }
    classes
}

/// The effective font colour: `(automatic, explicit colour)`.
///
/// ODF inherits `fo:color` and `style:use-window-font-color` separately
/// (§16.2) and an applying automatic flag wins (§20.395). LibreOffice instead
/// lets the nearest style that sets either decide (measured: an explicit span
/// colour over the automatic default paints; an automatic span over a
/// coloured parent is automatic). That reading is
/// `CompatRules::nearest_font_color`; without it, cascades where the two
/// readings differ reject.
fn font_color(cascade: &Cascade<'_>, rules: CompatRules) -> Result<(bool, Option<String>), String> {
    let key = |ns: &str, name: &str| (ns.to_string(), name.to_string());
    let (window, explicit) = (key(NS_STYLE, "use-window-font-color"), key(NS_FO, "color"));
    let mut nearest: Option<(bool, Option<String>)> = None;
    for style in &cascade.styles {
        let Some(properties) = style.group(cascade.group) else {
            continue;
        };
        let automatic = properties.attributes.get(&window).map(String::as_str) == Some("true");
        // Every applied style is checked for ambiguity, not only those
        // before the deciding one.
        match (automatic, properties.attributes.get(&explicit)) {
            (true, Some(_)) => return Err(unsupported("font-color-ambiguous")),
            (true, None) if nearest.is_none() => nearest = Some((true, None)),
            (false, Some(color)) if nearest.is_none() => {
                nearest = Some((false, Some(hex_color(color)?)))
            }
            _ => {}
        }
    }
    let nearest = nearest.unwrap_or((false, None));
    // Per-property inheritance: an automatic flag set anywhere applies.
    let per_property = if cascade.get(NS_STYLE, "use-window-font-color") == Some("true") {
        (true, None)
    } else {
        (
            false,
            cascade.get(NS_FO, "color").map(hex_color).transpose()?,
        )
    };
    if nearest == per_property || rules.nearest_font_color {
        Ok(nearest)
    } else {
        Err(unsupported("font-color-cascade"))
    }
}

/// A space no taller than its text: the same face scaled no larger (see
/// `emit_runs`).
fn space_fits_line(space: &RunProperties, glyph: &RunProperties) -> Result<(), String> {
    if space.family == glyph.family
        && space.bold == glyph.bold
        && space.italic == glyph.italic
        && space.position == glyph.position
        && space.size <= glyph.size
    {
        Ok(())
    } else {
        Err(unsupported("space-formatting"))
    }
}

/// The properties that set a line's height: face, size and raise/lower.
fn same_line_metrics(
    a: &RunProperties,
    b: &RunProperties,
    code: &'static str,
) -> Result<(), String> {
    if a.family == b.family && a.size == b.size && a.position == b.position {
        Ok(())
    } else {
        Err(unsupported(code))
    }
}

/// Whether a span's own style gives LibreOffice formatting for it: a
/// character style, or automatic properties other than a revision id.
fn span_formats(style: &Style) -> bool {
    !style.automatic
        || style.parent.is_some()
        || style.groups.get(TEXT).is_some_and(|p| {
            p.attributes
                .keys()
                .any(|key| (key.0.as_str(), key.1.as_str()) != (NS_OFFICEOOO, "rsid"))
        })
}

fn portion_character(c: char) -> bool {
    matches!(
        c,
        '\u{ad}' | '\u{a0}' | '\u{2011}' | '\u{200b}' | '\u{2060}'
    )
}

fn uncertain_portion_character(c: char) -> bool {
    matches!(c, '\u{feff}' | '\u{61c}' | '\u{200e}' | '\u{200f}' | '\u{202a}'..='\u{202e}' | '\u{2066}'..='\u{2069}')
}

/// The spans enclosing a context, outermost first: each span's character
/// style (an automatic style's parent) by display name, and its automatic
/// style's own text properties without revision ids. Spans that contribute
/// neither are left out (LibreOffice creates no formatting for them,
/// measured).
type SpanChain<'a> = Vec<(Option<&'a str>, BTreeMap<&'a crate::style::Key, &'a str>)>;

fn span_chain<'a>(contexts: &[TextContext<'a>], mut context: u32) -> SpanChain<'a> {
    let mut chain = Vec::new();
    while let Some(parent) = contexts[context as usize].parent {
        let styles = &contexts[context as usize].styles;
        let style = styles[0].style();
        let display = |style: &'a Style| {
            style
                .attributes
                .get(&(NS_STYLE.to_string(), "display-name".to_string()))
                .unwrap_or(&style.name)
                .as_str()
        };
        let entry = if style.automatic {
            let properties: BTreeMap<&crate::style::Key, &str> = style
                .groups
                .get(TEXT)
                .into_iter()
                .flat_map(|p| &p.attributes)
                .filter(|(key, _)| (key.0.as_str(), key.1.as_str()) != (NS_OFFICEOOO, "rsid"))
                .map(|(key, value)| (key, value.as_str()))
                .collect();
            (
                styles.get(1).map(|parent| display(parent.style())),
                properties,
            )
        } else {
            (Some(display(style)), BTreeMap::new())
        };
        if entry.0.is_some() || !entry.1.is_empty() {
            chain.push(entry);
        }
        context = parent;
    }
    chain.reverse();
    chain
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum Seam {
    Same,
    Boundary,
    Unknown,
}

#[derive(Default)]
struct SeamCache<'a> {
    chains: HashMap<u32, SpanChain<'a>>,
    /// Resolved Latin formatting and Asian font size.
    resolved: HashMap<u32, (RunProperties, f64)>,
    seams: HashMap<(u32, u32), Seam>,
}

/// §6.1.2 white-space collapsing over a paragraph's character data: tab, CR
/// and LF become spaces, leading and trailing spaces are removed and runs of
/// spaces collapse to one. `text:s`, `text:tab` and `text:line-break` keep
/// the characters they represent and are not collapsed.
///
/// A portion boundary (`Token::Boundary`) is not character data: spaces on
/// both sides of it still collapse, and it follows the space that is kept.
fn collapse(tokens: Vec<Token>) -> (Vec<Token>, Option<u32>) {
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
            Token::Char(' ' | '\t' | '\r' | '\n', context) => match pending {
                None if started => pending = Some(Token::Char(' ', context)),
                // A space of other formatting merges into the kept one: where
                // LibreOffice ends its portion there is unknown.
                Some(Token::Char(_, kept)) if kept != context => {
                    boundaries.push(Token::Boundary(false));
                }
                _ => {}
            },
            Token::Char(c, context) => {
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                out.append(&mut boundaries);
                started = true;
                out.push(Token::Char(c, context));
            }
            element => {
                // Spaces adjacent to an element are character data on either
                // side of it; only consecutive character-data spaces merge.
                if let Some(space) = pending.take() {
                    out.push(space);
                }
                out.append(&mut boundaries);
                started = true;
                out.push(element);
            }
        }
    }
    // A trailing collapsed space is removed (step 5); its context is
    // returned for the LibreOffice comparison in `emit_runs`.
    let removed = match pending {
        Some(Token::Char(_, context)) => Some(context),
        _ => None,
    };
    (out, removed)
}

/// Effective text properties of one script class.
#[derive(Clone, PartialEq, Debug)]
struct RunProperties {
    family: Atom,
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
    lang_east_asia: Option<Atom>,
    lang_bidi: Option<Atom>,
}

impl RunProperties {
    fn resolve(
        cascade: &Cascade<'_>,
        class: ScriptClass,
        fonts: &FontFaces,
        rules: CompatRules,
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
            // A positiveLength (§20.190). LibreOffice renders sizes below
            // 1 pt as 1 pt and lays out sizes from about 3,000 pt differently
            // (measured); neither behaviour is adopted, so the admitted
            // domain is 1-1,000 pt (measured equal up to 2,800 pt).
            Some(value) => match length_pt(value)? {
                size if (1.0..=1000.0).contains(&size) => size,
                _ => return Err(unsupported("font-size-range")),
            },
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
        let (mut color_auto, mut color) = font_color(cascade, rules)?;
        // Character opacity of the font colour (LibreOffice extension):
        // LibreOffice writes 100% with an explicit colour and 0% with the
        // automatic colour, where it has no effect (measured); only those
        // pairs are inert.
        match cascade.get(NS_LOEXT, "opacity") {
            None | Some("100%") => {}
            Some("0%") if color_auto => {}
            Some(_) => return Err(unsupported("text-opacity")),
        }
        // §20.395 automatic (or unset) colour over the run's own shading,
        // resolved as LibreOffice paints it.
        if let (None, Some(background)) = (&color, &background) {
            color = Some(libreoffice_automatic_color(background)?.to_string());
            color_auto = false;
        }
        // Letter spacing: LibreOffice spaces between the characters of a
        // portion, the DOCX layout after each character with an advance
        // floor (measured: -1 pt over four letters moves the next word 3 pt
        // and 4 pt). Rejected until mapped.
        let letter_spacing: Option<f64> = match cascade.get(NS_FO, "letter-spacing") {
            None | Some("normal") => None,
            Some(value) if length_pt(value)? == 0.0 => None,
            Some(_) => return Err(unsupported("letter-spacing")),
        };
        // Horizontal scaling is not yet measured against LibreOffice (and
        // non-positive values break the DOCX layout); only 100% is admitted.
        let scale: Option<f64> = match cascade.get(NS_STYLE, "text-scale") {
            None => None,
            Some(value) if percent(value)? == 1.0 => None,
            Some(_) => return Err(unsupported("text-scale")),
        };
        // §20.384: a raised/lowered or resized portion changes the line
        // height differently in LibreOffice and the DOCX layout (measured: a
        // -20% 58% portion and a 0% 0% portion); only the inert value
        // (no raise, 100% size) is admitted until the geometry is measured.
        match cascade
            .get(NS_STYLE, "text-position")
            .map(str::split_whitespace)
        {
            None => {}
            Some(mut parts) => {
                // Keywords (`super`, `sub`) are not percentages: rejected too.
                let raise = parts.next().and_then(|v| percent(v).ok());
                let relative = parts.next().and_then(|v| percent(v).ok());
                if raise != Some(0.0) || relative != Some(1.0) || parts.next().is_some() {
                    return Err(unsupported("text-position"));
                }
            }
        }
        let position: Option<f64> = None;
        // §20.227 follows XSL's language-sensitive case mapping; the DOCX
        // renderer uppercases locale-independently. Not yet measured.
        if cascade
            .get(NS_FO, "text-transform")
            .is_some_and(|v| v != "none")
        {
            return Err(unsupported("text-transform"));
        }
        let caps = false;
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

    /// Native size of one emitted run, excluding its text (charged with the
    /// retained text).
    fn retained_estimate(&self) -> usize {
        let strings = 4 * self.family.len()
            + self.underline.map_or(0, str::len)
            + [&self.color, &self.background]
                .iter()
                .map(|value| value.as_ref().map_or(0, String::len))
                .sum::<usize>()
            + [&self.lang_east_asia, &self.lang_bidi]
                .iter()
                .map(|value| value.as_ref().map_or(0, |atom| atom.len()))
                .sum::<usize>();
        std::mem::size_of::<DocRun>() + std::mem::size_of::<TextRun>() + strings
    }

    fn text_run(&self) -> TextRun {
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
            lang_east_asia: self.lang_east_asia.as_deref().map(str::to_string),
            lang_bidi: self.lang_bidi.as_deref().map(str::to_string),
            ..TextRun::default()
        }
    }

    fn font_facts(&self) -> RunFontFacts {
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
                    || (n.starts_with("font-size") && !n.starts_with("font-size-rel"))
                    || n.starts_with("font-weight")
                    || (n.starts_with("font-style") && !n.starts_with("font-style-name")) =>
            {
                true
            }
            (
                NS_FO,
                "font-family" | "font-size" | "font-weight" | "font-style" | "color"
                | "background-color" | "letter-spacing" | "text-transform" | "font-variant"
                | "script" | "hyphenate",
            ) => true,
            // §20.202 western language: admitted but not carried (the DOCX
            // model has no western slot). Language-specific OpenType forms
            // are not applied, as for DOCX w:lang: #2001.
            (NS_FO, "language" | "country") => true,
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
            // Validated against the resolved colour in `RunProperties::resolve`.
            (NS_LOEXT, "opacity") => true,
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
                family: fonts.atom(&unquote_family(family)?),
                generic: generic_class(attribute("font-family-generic").map(String::as_str))?,
                pitch: pitch(attribute("font-pitch").map(String::as_str))?,
                charset: charset(attribute("font-charset").map(String::as_str))?,
            });
        }
    }
    Err(unsupported("font-unresolved"))
}

/// The colour LibreOffice paints automatic text in over an opaque
/// `background` (hex rrggbb): white when `Color::IsDark()`, else black
/// (`SwDrawTextInfo::ApplyAutoColor`, sw/source/core/txtnode/fntcache.cxx;
/// the PDF export takes the same path). `IsDark()`
/// (tools/source/generic/color.cxx, LibreOffice 25.2 to 26.8): the WCAG
/// relative luminance scaled to 255 and truncated, at most 87; LibreOffice's
/// default shape fill #729fcf instead uses the integer luminance
/// (76 R + 151 G + 29 B) >> 8, at most 62. Measured identical on 953
/// backgrounds (every grey, the standard palette, channel ramps and random
/// colours). The DOCX automatic colour uses another threshold, so the
/// reader writes the resolved colour.
fn libreoffice_automatic_color(background: &str) -> Result<&'static str, String> {
    let channel = |at: usize| {
        background
            .get(at..at + 2)
            .and_then(|hex| u8::from_str_radix(hex, 16).ok())
            .ok_or_else(|| invalid("a colour is not #rrggbb"))
    };
    let (r, g, b) = (channel(0)?, channel(2)?, channel(4)?);
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
        // C++ double-to-integer conversion truncates.
        (luminance * 255.0) as u8 <= 87
    };
    Ok(if dark { "ffffff" } else { "000000" })
}

/// LibreOffice writes its area fill (`loext:graphic-properties`) beside
/// `fo:background-color`: `draw:fill="none"` with a transparent background,
/// or `draw:fill="solid"` repeating the background colour. Those pairs are
/// the paragraph background (§20.182; `draw:fill` §20.118, `draw:fill-color`
/// §20.119, `draw:opacity` §20.152); a fill that disagrees, other fill
/// kinds (gradient, bitmap, hatch) and partial opacity reject.
fn check_paragraph_fill(paragraph: &Cascade<'_>, fill: &Cascade<'_>) -> Result<(), String> {
    if !fill.element_names().is_empty() {
        return Err(unsupported("paragraph-fill"));
    }
    for (ns, name) in fill.keys() {
        if !matches!(
            (ns.as_str(), name.as_str()),
            (NS_DRAW, "fill" | "fill-color" | "opacity")
        ) {
            return Err(unsupported("paragraph-fill"));
        }
    }
    let background = match paragraph.get(NS_FO, "background-color") {
        None | Some("transparent") => None,
        Some(value) => Some(hex_color(value)?),
    };
    let agrees = match fill.get(NS_DRAW, "fill") {
        None if fill.keys().is_empty() => true,
        Some("none") => background.is_none(),
        Some("solid") => {
            let color = fill.get(NS_DRAW, "fill-color").map(hex_color).transpose()?;
            background.is_some()
                && color == background
                && match fill.get(NS_DRAW, "opacity") {
                    None => true,
                    Some(value) => percent(value)? == 1.0,
                }
        }
        _ => false,
    };
    if agrees {
        Ok(())
    } else {
        Err(unsupported("paragraph-fill"))
    }
}

/// Paragraph properties: exact mappings, inert values, rejection otherwise.
fn paragraph_properties(cascade: &Cascade<'_>, rules: CompatRules) -> Result<Collected, String> {
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
            // LibreOffice 26 writes these on every default style: greedy
            // line breaking (as the DOCX layout breaks lines) and no extra
            // tab distance.
            (NS_LOEXT, "paragraph-composer") => value == "false",
            (NS_LOEXT, "tab-stop-distance") => length_pt(value).is_ok_and(|v| v == 0.0),
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
            // A solid colour maps to the DOCX paragraph shading (below).
            (NS_FO, "background-color") => value == "transparent" || hex_color(value).is_ok(),
            (NS_FO, "border" | "border-top" | "border-bottom" | "border-left" | "border-right") => {
                value == "none"
            }
            // Padding insets the text even without a border (§20.217-§20.221).
            (NS_FO, n) if n.starts_with("padding") => length_pt(value).is_ok_and(|v| v == 0.0),
            // Border joining is inert without borders.
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
        // Proportional spacing. ODF 1.4 §20.204 defines the percentage on the
        // font size; LibreOffice (which writes 1.4) scales the natural line
        // height instead — DOCX "auto" with the percentage as the multiple,
        // measured equal at 100% and 150% on uniform text. That is an
        // user-approved LibreOffice compatibility rule (2026-10-31) on text
        // of one face and size (checked per paragraph). Line pitch matches
        // from 50% to 300%, but below 100% LibreOffice shrinks the line from
        // the top, setting the glyphs higher than the DOCX layout (measured:
        // 3 pt at 50%, 1.5 pt at 75%); the admitted scope is 100-300% and
        // everything else rejects.
        (Some(value), None) if value.ends_with('%') => {
            let multiple = percent(value)?;
            if !rules.proportional_line_height || !(1.0..=3.0).contains(&multiple) {
                return Err(unsupported("line-height-proportional"));
            }
            Some(LineSpacing {
                value: multiple,
                rule: "auto".into(),
                explicit: true,
            })
        }
        // Fixed and minimum line heights: LibreOffice places the text at the
        // bottom of a taller line box (all extra space above), measured 7.5pt
        // lower than the DOCX layout for a 30pt minimum on 12pt text and
        // 1.5pt lower for a fixed 20pt line. DOCX has no control for that
        // placement, so both are rejected until a model capability exists.
        (Some(_), None) => return Err(unsupported("line-height-fixed")),
        (None, Some(_)) => return Err(unsupported("line-height-at-least")),
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
        shading: match cascade.get(NS_FO, "background-color") {
            None | Some("transparent") => None,
            Some(value) => Some(hex_color(value)?),
        },
        ..DocParagraph::default()
    };
    Ok(Collected {
        paragraph,
        top: length("margin-top")?,
        bottom: length("margin-bottom")?,
        contextual: cascade.get(NS_STYLE, "contextual-spacing") == Some("true"),
        named_style: None,
        area: 0,
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

    const NS: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:officeooo="http://openoffice.org/2009/office""#;
    const FACES: &str = r#"<office:font-face-decls><style:font-face style:name="Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman"/><style:font-face style:name="Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss"/><style:font-face style:name="Sans2" svg:font-family="'Liberation Sans'" style:font-family-generic="system"/><style:font-face style:name="Mincho" svg:font-family="'Hiragino Mincho ProN'" style:font-family-generic="system"/></office:font-face-decls>"#;

    struct Fixture<'a> {
        default_text: &'a str,
        /// Additional `style:font-face` declarations in styles.xml.
        faces: &'a str,
        /// Additional `style:page-layout-properties` attributes.
        page: &'a str,
        /// Additional default paragraph style `style:paragraph-properties`
        /// attributes.
        default_paragraph: &'a str,
        styles: &'a str,
        automatic: &'a str,
        master: &'a str,
        body: &'a str,
        /// `config:config-item` elements for settings.xml (none: no part).
        settings: &'a str,
    }

    impl Default for Fixture<'_> {
        fn default() -> Self {
            Self {
                default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Serif" style:font-size-asian="12pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true""#,
                faces: "",
                page: "",
                default_paragraph: "",
                styles: "",
                automatic: "",
                master: "",
                body: "<text:p>x</text:p>",
                settings: "",
            }
        }
    }

    /// The fixture page geometry with `extra` attributes added or replacing
    /// the defaults (` name="value"` pairs).
    fn page_properties(extra: &str) -> String {
        let mut attributes: Vec<(String, String)> = [
            ("fo:page-width", "210mm"),
            ("fo:page-height", "297mm"),
            ("fo:margin-top", "20mm"),
            ("fo:margin-bottom", "20mm"),
            ("fo:margin-left", "25mm"),
            ("fo:margin-right", "25mm"),
        ]
        .iter()
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect();
        for pair in extra.split('"').collect::<Vec<_>>().chunks(2) {
            if let [name, value] = pair {
                let name = name.trim().trim_end_matches('=').to_string();
                match attributes.iter_mut().find(|(k, _)| *k == name) {
                    Some(entry) => entry.1 = value.to_string(),
                    None => attributes.push((name, value.to_string())),
                }
            }
        }
        attributes
            .iter()
            .map(|(k, v)| format!(r#" {k}="{v}""#))
            .collect()
    }

    fn read_fixture(fixture: Fixture<'_>) -> Result<Document, String> {
        read_fixture_with(fixture, LIBREOFFICE_RULES)
    }

    /// The reader without the LibreOffice compatibility rules: the spec-only
    /// behaviour each rule replaces.
    const SPEC_ONLY: CompatRules = CompatRules {
        proportional_line_height: false,
        weak_script_class: false,
        nearest_font_color: false,
    };

    fn read_fixture_with(fixture: Fixture<'_>, rules: CompatRules) -> Result<Document, String> {
        let styles = format!(
            r#"<office:document-styles {NS} office:version="1.3">{}<office:styles><style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.25cm"{}/><style:text-properties {}/></style:default-style><style:default-style style:family="graphic"/>{}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties{}/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">{}</style:master-page></office:master-styles></office:document-styles>"#,
            FACES.replace(
                "</office:font-face-decls>",
                &format!("{}</office:font-face-decls>", fixture.faces)
            ),
            fixture.default_paragraph,
            fixture.default_text,
            fixture.styles,
            page_properties(fixture.page),
            fixture.master
        );
        let content = format!(
            r#"<office:document-content {NS} office:version="1.3">{FACES}<office:automatic-styles>{}</office:automatic-styles><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            fixture.automatic, fixture.body
        );
        let settings = format!(
            r#"<office:document-settings {NS} xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings">{}</config:config-item-set></office:settings></office:document-settings>"#,
            fixture.settings
        );
        let mut files = vec![
            ("META-INF/manifest.xml", manifest(ODT, "")),
            ("styles.xml", styles),
            ("content.xml", content),
        ];
        if !fixture.settings.is_empty() {
            files.push(("settings.xml", settings));
        }
        let files: Vec<(&str, &str)> = files
            .iter()
            .map(|(name, body)| (*name, body.as_str()))
            .collect();
        let bytes = package(ODT, &files);
        let package = OdfPackage::open(
            bytes,
            OoxmlFormat::Docx,
            PackageLimits::default(),
            &MEDIA_TYPES,
        )?;
        read_with(&package, rules)
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
    fn collapses_white_space_and_keeps_explicit_spaces_and_breaks() {
        let document = read_fixture(Fixture {
            body: "<text:p>\n  a \t b<text:s text:c=\"2\"/>cd<text:line-break/>  e  </text:p>",
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(texts(paragraphs(&document)[0]), ["a b  cd", "<br>", " e"]);
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
        let gaps: Vec<f64> = paragraphs(&document)
            .iter()
            .map(|p| p.space_before)
            .collect();
        assert_eq!(gaps, [6.0, 16.0, 22.0, 0.0, 18.0]);
        assert!(paragraphs(&document)
            .iter()
            .all(|p| p.space_after == 0.0 && !p.contextual_spacing));

        let mixed = read_fixture(Fixture {
            styles,
            automatic: r#"<style:style style:name="P1" style:family="paragraph" style:parent-style-name="C"><style:paragraph-properties fo:margin-top="12pt" fo:margin-bottom="12pt" style:contextual-spacing="false"/></style:style>"#,
            body: r#"<text:p text:style-name="C">1</text:p><text:p text:style-name="P1">2</text:p>"#,
            ..Fixture::default()
        });
        assert_eq!(mixed.unwrap_err(), unsupported("contextual-spacing-mixed"));
    }

    #[test]
    fn a_hard_page_break_carries_only_the_new_top_margin() {
        let document = read_fixture(Fixture {
            styles: r#"<style:style style:name="A" style:family="paragraph"><style:paragraph-properties fo:margin-top="6pt" fo:margin-bottom="10pt"/></style:style>"#,
            automatic: r#"<style:style style:name="Before" style:family="paragraph" style:parent-style-name="A"><style:paragraph-properties fo:break-before="page"/></style:style><style:style style:name="After" style:family="paragraph" style:parent-style-name="A"><style:paragraph-properties fo:break-after="page"/></style:style>"#,
            body: r#"<text:p text:style-name="A">1</text:p><text:p text:style-name="Before">2</text:p><text:p text:style-name="After">3</text:p><text:p text:style-name="A">4</text:p><text:p text:style-name="A">5</text:p>"#,
            ..Fixture::default()
        })
        .unwrap();
        let spacing: Vec<(f64, bool)> = paragraphs(&document)
            .iter()
            .map(|p| (p.space_before, p.page_break_before))
            .collect();
        assert_eq!(
            spacing,
            [
                (6.0, false),
                (6.0, true),
                (16.0, false),
                (6.0, true),
                (16.0, false)
            ]
        );
    }

    #[test]
    fn maps_line_height_controls_and_rejects_their_combination() {
        let style = |props: &str| {
            format!(
                r#"<style:style style:name="L" style:family="paragraph"><style:paragraph-properties {props}/></style:style>"#
            )
        };
        let spacing = |props: &str| {
            read_fixture(Fixture {
                styles: &style(props),
                body: r#"<text:p text:style-name="L">x</text:p>"#,
                ..Fixture::default()
            })
            .map(|d| {
                paragraphs(&d)[0]
                    .line_spacing
                    .as_ref()
                    .map(|s| (s.value, s.rule.clone()))
            })
        };
        // User-approved LibreOffice compatibility rule (2026-10-31): the
        // admitted 100-300% scope maps to DOCX auto; outside it rejects.
        for (height, multiple) in [("115%", 1.15), ("100%", 1.0), ("300%", 3.0)] {
            assert_eq!(
                spacing(&format!(r#"fo:line-height="{height}""#)).unwrap(),
                Some((multiple, "auto".to_string())),
                "{height}"
            );
        }
        for height in ["50%", "99.9%", "300.1%"] {
            assert_eq!(
                spacing(&format!(r#"fo:line-height="{height}""#)).unwrap_err(),
                unsupported("line-height-proportional"),
                "{height}"
            );
        }
        assert_eq!(
            spacing(r#"fo:line-height="20pt""#).unwrap_err(),
            unsupported("line-height-fixed")
        );
        assert_eq!(
            spacing(r#"style:line-height-at-least="30pt""#).unwrap_err(),
            unsupported("line-height-at-least")
        );
        assert_eq!(
            spacing(r#"fo:line-height="normal" style:line-height-at-least="30pt""#).unwrap(),
            None
        );
        assert_eq!(
            spacing(r#"fo:line-height="20pt" style:line-height-at-least="30pt""#).unwrap_err(),
            unsupported("line-height-combination")
        );
        assert_eq!(
            spacing(r#"style:line-spacing="2pt""#).unwrap_err(),
            unsupported("paragraph-property")
        );
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
            .map(|r| match r {
                DocRun::Text(t) => t.as_ref(),
                _ => panic!(),
            })
            .collect();
        assert_eq!(
            (
                runs[0].font_size,
                runs[0].bold,
                runs[0].font_family.as_deref()
            ),
            (20.0, false, Some("Liberation Serif"))
        );
        assert_eq!(
            (runs[1].bold, runs[1].italic, runs[1].font_family.as_deref()),
            (true, false, Some("Liberation Sans"))
        );
        assert_eq!(
            (runs[2].bold, runs[2].italic, runs[2].font_size),
            (true, true, 20.0)
        );
        // Only used faces enter the font table; the unused "Sans2" twin with a
        // different generic class does not conflict.
        assert_eq!(
            document
                .font_family_classes
                .get("Liberation Sans")
                .map(String::as_str),
            Some("swiss")
        );
    }

    #[test]
    fn rejects_unadmitted_content_with_reason_codes() {
        let reject = |fixture: Fixture<'_>| read_fixture(fixture).unwrap_err();
        assert_eq!(
            reject(Fixture {
                body: "<text:table-of-content/>",
                ..Fixture::default()
            }),
            unsupported("body-element")
        );
        assert_eq!(
            reject(Fixture {
                body: "<text:p>日本</text:p>",
                ..Fixture::default()
            }),
            unsupported("script-class")
        );
        assert_eq!(
            reject(Fixture {
                master: "<style:header><text:p>h</text:p></style:header>",
                ..Fixture::default()
            }),
            unsupported("header-footer")
        );
        assert_eq!(
            reject(Fixture {
                default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:letter-kerning="false""#,
                ..Fixture::default()
            }),
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
                body: "<text:p>\u{0500}x</text:p>",
                ..Fixture::default()
            }),
            unsupported("script-unmapped-character")
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
    fn long_paragraphs_merge_runs_linearly_and_respect_the_character_bound() {
        let text = "word ".repeat(200_000);
        let start = std::time::Instant::now();
        let document = read_fixture(Fixture {
            body: &format!("<text:p>{text}</text:p>"),
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(paragraphs(&document)[0].runs.len(), 1);
        // A generous bound that the former quadratic merge exceeded by orders
        // of magnitude; it only guards against regressions in complexity.
        assert!(start.elapsed().as_secs() < 20, "{:?}", start.elapsed());
        let spaces = read_fixture(Fixture {
            body: &format!(
                "<text:p>{}</text:p>",
                r#"<text:s text:c="600000"/>"#.repeat(2)
            ),
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(spaces.contains("character limit"));
    }

    #[test]
    fn characters_outside_table_22_need_class_independent_properties() {
        let document = read_fixture(Fixture {
            body: "<text:p>a\u{2014}b c</text:p>",
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(texts(paragraphs(&document)[0]), ["a\u{2014}b c"]);
        let error = read_fixture(Fixture {
            default_text: r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="30pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true""#,
            body: "<text:p>a\u{0500}b</text:p>",
            ..Fixture::default()
        })
        .unwrap_err();
        assert_eq!(error, unsupported("script-unmapped-character"));
    }

    #[test]
    fn rejects_rendering_configuration_the_reader_does_not_apply() {
        let reject = |fixture: Fixture<'_>| read_fixture(fixture).unwrap_err();
        assert_eq!(
            reject(Fixture {
                styles: "<text:linenumbering-configuration/>",
                ..Fixture::default()
            }),
            unsupported("line-numbering")
        );
        assert_eq!(
            reject(Fixture {
                styles: r#"<style:default-style style:family="text"><style:text-properties fo:font-style="italic"/></style:default-style>"#,
                ..Fixture::default()
            }),
            unsupported("text-default-style")
        );
        assert_eq!(
            reject(Fixture {
                styles: r#"<style:style style:name="H" style:family="paragraph"><style:map style:condition="outline-level()=1" style:apply-style-name="Standard"/></style:style>"#,
                body: r#"<text:p text:style-name="H">x</text:p>"#,
                ..Fixture::default()
            }),
            unsupported("conditional-style")
        );
        assert_eq!(
            reject(Fixture {
                styles: r#"<style:style style:name="R" style:family="paragraph"><style:text-properties style:font-size-rel="8pt"/></style:style>"#,
                body: r#"<text:p text:style-name="R">x</text:p>"#,
                ..Fixture::default()
            }),
            unsupported("text-property")
        );
        assert!(check_settings(r#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="ParaSpaceMax" config:type="boolean">true</config:config-item></config:config-item-set></office:settings></office:document-settings>"#, Some("1.3")).is_err());
        assert!(check_settings(r#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><config:config-item config:name="ParaSpaceMax" config:type="boolean">false</config:config-item></office:document-settings>"#, Some("1.3")).is_ok());
    }

    #[test]
    fn nested_spans_share_style_chains_and_style_resolution_is_bounded() {
        // A 64-deep text style chain (five attributes per style) under 100
        // leaf styles: each span adds a 65-style chain to its §16.2 search
        // order.
        let mut styles = String::new();
        for i in 0..64 {
            let parent = if i == 0 {
                String::new()
            } else {
                format!(r#" style:parent-style-name="T{}""#, i - 1)
            };
            styles.push_str(&format!(
                r#"<style:style style:name="T{i}" style:family="text"{parent}><style:text-properties fo:font-size="{}pt" fo:hyphenation-remain-char-count="2" fo:hyphenation-push-char-count="2" fo:hyphenation-ladder-count="no-limit" fo:hyphenation-keep="auto"/></style:style>"#,
                10 + i % 4
            ));
        }
        for i in 0..100 {
            styles.push_str(&format!(
                r#"<style:style style:name="A{i}" style:family="text" style:parent-style-name="T63"/>"#
            ));
        }
        let nest = |leaves: &[usize]| {
            let open: String = leaves
                .iter()
                .map(|leaf| format!(r#"<text:span text:style-name="A{leaf}">"#))
                .collect();
            format!("{open}x{}", "</text:span>".repeat(leaves.len()))
        };
        let read_body = |body: &str| {
            read_fixture(Fixture {
                styles: &styles,
                body,
                ..Fixture::default()
            })
        };

        // Repeated spans under the same parent reuse one context.
        let repeated = format!("<text:p>{}</text:p>", nest(&[1, 2, 3]).repeat(500));
        let document = read_body(&repeated).unwrap();
        assert_eq!(texts(paragraphs(&document)[0]).concat(), "x".repeat(500));

        // A search order beyond the implementation bound.
        let deep: Vec<usize> = (0..32).collect();
        assert!(read_body(&format!("<text:p>{}</text:p>", nest(&deep)))
            .unwrap_err()
            .contains("style search order limit"));

        // Resolution work is bounded document-wide: each paragraph resolves
        // its (identical) three-span context again, ~1,170 units each.
        let paragraph = format!("<text:p>{}</text:p>", nest(&[1, 2, 3]));
        let error = read_body(&paragraph.repeat(15_000)).unwrap_err();
        assert!(error.contains("style resolution limit"), "{error}");
    }

    #[test]
    fn attribute_heavy_styles_and_many_faces_stay_bounded() {
        // Hyphenation parameters are admitted by prefix (inert while
        // hyphenation is off), so one style can carry many of them; every
        // paragraph resolution is charged for each.
        let heavy: String = (0..200)
            .map(|i| format!(r#" fo:hyphenation-x{i}="1""#))
            .collect();
        let styles = format!(
            r#"<style:style style:name="H" style:family="paragraph"><style:paragraph-properties{heavy}/></style:style>"#
        );
        let body = r#"<text:p text:style-name="H">x</text:p>"#.repeat(9000);
        let error = read_fixture(Fixture {
            styles: &styles,
            body: &body,
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(error.contains("style resolution limit"), "{error}");

        // Many distinct faces, each used by its own span: the used-face table
        // is a hash set.
        let faces = 5000;
        let mut styles = String::new();
        let mut body = String::from("<text:p>");
        for i in 0..faces {
            styles.push_str(&format!(
                r#"<style:style style:name="F{i}" style:family="text"><style:text-properties style:font-name="Face{i}" style:font-name-asian="Face{i}" style:font-name-complex="Face{i}"/></style:style>"#
            ));
            body.push_str(&format!(
                r#"<text:span text:style-name="F{i}">x</text:span>"#
            ));
        }
        body.push_str("</text:p>");
        let decls: String = (0..faces)
            .map(|i| {
                format!(r#"<style:font-face style:name="Face{i}" svg:font-family="'Face {i}'"/>"#)
            })
            .collect();
        let document = read_fixture(Fixture {
            faces: &decls,
            styles: &styles,
            body: &body,
            ..Fixture::default()
        });
        let document = document.unwrap();
        assert_eq!(paragraphs(&document)[0].runs.len(), faces);
    }

    #[test]
    fn retained_runs_are_charged_against_the_construction_limit() {
        // Alternating span styles make every character its own run.
        let styles = r#"<style:style style:name="A" style:family="text"><style:text-properties fo:font-size="11pt"/></style:style><style:style style:name="B" style:family="text"><style:text-properties fo:font-size="13pt"/></style:style>"#;
        let per_run = std::mem::size_of::<DocRun>() + std::mem::size_of::<TextRun>();
        let runs = MAX_CONSTRUCTION_BYTES / per_run + 1;
        let pair = r#"<text:span text:style-name="A">x</text:span><text:span text:style-name="B">x</text:span>"#;
        let paragraph = format!("<text:p>{}</text:p>", pair.repeat(5000));
        let body = paragraph.repeat(runs / 10_000 + 1);
        let error = read_fixture(Fixture {
            styles,
            body: &body,
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(error.contains("construction limit"), "{error}");
    }

    #[test]
    fn shared_property_strings_and_line_breaks_stay_bounded() {
        // Many distinct span contexts resolve to one long family: the face's
        // interned family is shared, and equal properties merge into one run.
        let family = "F".repeat(100_000);
        let faces = format!(r#"<style:font-face style:name="Long" svg:font-family="'{family}'"/>"#);
        let styles: String = (0..2000)
            .map(|i| format!(r#"<style:style style:name="S{i}" style:family="text"><style:text-properties style:font-name="Long" style:font-name-asian="Long" style:font-name-complex="Long"/></style:style>"#))
            .collect();
        let body = format!(
            "<text:p>{}</text:p>",
            (0..2000)
                .map(|i| format!(r#"<text:span text:style-name="S{i}">x</text:span>"#))
                .collect::<String>()
        );
        let document = read_fixture(Fixture {
            faces: &faces,
            styles: &styles,
            body: &body,
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(texts(paragraphs(&document)[0]), vec!["x".repeat(2000)]);

        // Every line break is a retained run and costs downstream layout work;
        // either bound stops a break flood.
        let breaks = MAX_CONSTRUCTION_BYTES / std::mem::size_of::<DocRun>() + 1;
        let paragraph = format!(
            "<text:p>x{}</text:p>",
            "<text:line-break/>x".repeat(400_000)
        );
        let error = read_fixture(Fixture {
            body: &paragraph.repeat(breaks / 400_000 + 1),
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(
            error.contains("construction limit") || error.contains("line break layout limit"),
            "{error}"
        );
    }

    #[test]
    fn rejects_an_active_layout_grid() {
        assert_eq!(
            read_fixture(Fixture {
                page: r#" style:layout-grid-mode="line" style:layout-grid-lines="20""#,
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("page-layout-property")
        );
        assert!(read_fixture(Fixture {
            page: r#" style:layout-grid-mode="none" style:layout-grid-lines="20""#,
            ..Fixture::default()
        })
        .is_ok());
    }

    #[test]
    fn style_identifiers_and_expanded_names_are_charged() {
        // A namespace URI expanded into every stored attribute key of many
        // unused styles.
        let namespace = format!("urn:x:{}", "n".repeat(4000));
        let unused: String = (0..8000)
            .map(|i| format!(r#"<style:style style:name="U{i}" style:family="paragraph" xmlns:x="{namespace}" x:a="1" x:b="1"/>"#))
            .collect();
        let error = read_fixture(Fixture {
            styles: &unused,
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(error.contains("style storage limit"), "{error}");

        // A long parent name is looked up once per style, not per paragraph.
        let parent = "P".repeat(1_000_000);
        let styles = format!(
            r#"<style:style style:name="{parent}" style:family="paragraph"/><style:style style:name="C" style:family="paragraph" style:parent-style-name="{parent}"/>"#
        );
        let body = r#"<text:p text:style-name="C">x</text:p>"#.repeat(3000);
        let document = read_fixture(Fixture {
            styles: &styles,
            body: &body,
            ..Fixture::default()
        })
        .unwrap();
        assert_eq!(paragraphs(&document).len(), 3000);
    }

    #[test]
    fn line_breaks_must_share_the_line_metrics_docx_uses() {
        let styles = r#"<style:style style:name="Big" style:family="text"><style:text-properties fo:font-size="30pt"/></style:style><style:style style:name="Bold" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>"#;
        let read = |body: &str| {
            read_fixture(Fixture {
                styles,
                body,
                ..Fixture::default()
            })
        };
        // Same face and size as the preceding text (boldness is irrelevant).
        assert!(read(
            r#"<text:p>A<text:span text:style-name="Bold"><text:line-break/></text:span>B</text:p>"#
        )
        .is_ok());
        assert_eq!(
            read(r#"<text:p>A<text:span text:style-name="Big"><text:line-break/></text:span>B</text:p>"#).unwrap_err(),
            unsupported("line-break-formatting")
        );
        // A line holding only a break: leading, consecutive or trailing.
        for body in [
            "<text:p><text:line-break/>B</text:p>",
            "<text:p>A<text:line-break/><text:line-break/>B</text:p>",
            "<text:p>A<text:line-break/></text:p>",
            "<text:p><text:line-break/></text:p>",
        ] {
            assert_eq!(
                read(body).unwrap_err(),
                unsupported("line-break-empty-line"),
                "{body}"
            );
        }
    }

    #[test]
    fn consecutive_line_breaks_are_charged_for_downstream_layout() {
        // 12,000 breaks between text cost ~144M units in the DOCX line
        // breaker's run-index search.
        let body = format!("<text:p>x{}</text:p>", "<text:line-break/>x".repeat(12_000));
        let error = read_fixture(Fixture {
            body: &body,
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(error.contains("line break layout limit"), "{error}");
        let body = format!("<text:p>x{}</text:p>", "<text:line-break/>x".repeat(3_000));
        assert!(read_fixture(Fixture {
            body: &body,
            ..Fixture::default()
        })
        .is_ok());
    }

    #[test]
    fn the_font_table_is_charged_against_the_construction_limit() {
        // Six 6 MiB families: their runs (four copies each) fit, but the font
        // table and its metadata maps copy every used family four more times.
        let families: Vec<String> = (0..6)
            .map(|i| format!("{i}{}", "F".repeat(6 << 20)))
            .collect();
        let faces: String = families
            .iter()
            .enumerate()
            .map(|(i, family)| format!(r#"<style:font-face style:name="L{i}" svg:font-family="'{family}'" style:font-family-generic="roman"/>"#))
            .collect();
        let styles: String = (0..6)
            .map(|i| format!(r#"<style:style style:name="S{i}" style:family="text"><style:text-properties style:font-name="L{i}" style:font-name-asian="L{i}" style:font-name-complex="L{i}"/></style:style>"#))
            .collect();
        let body: String = (0..6)
            .map(|i| format!(r#"<text:p><text:span text:style-name="S{i}">x</text:span></text:p>"#))
            .collect();
        let error = read_fixture(Fixture {
            faces: &faces,
            styles: &styles,
            body: &body,
            ..Fixture::default()
        })
        .unwrap_err();
        assert!(error.contains("construction limit"), "{error}");
    }

    #[test]
    fn tabs_reject_and_lines_around_breaks_need_glyphs() {
        let read = |body: &str| {
            read_fixture(Fixture {
                body,
                ..Fixture::default()
            })
        };
        assert_eq!(
            read("<text:p>a<text:tab/>b</text:p>").unwrap_err(),
            unsupported("tab")
        );
        // A line between breaks holding only white space.
        assert_eq!(
            read("<text:p>A<text:line-break/><text:s/><text:line-break/>B</text:p>").unwrap_err(),
            unsupported("line-break-empty-line")
        );
        assert!(read("<text:p>A<text:line-break/>x<text:line-break/>B</text:p>").is_ok());
    }

    #[test]
    fn rejects_raised_lowered_or_resized_text() {
        for value in ["-20% 58%", "0% 58%", "super 58%", "33%"] {
            let styles = format!(
                r#"<style:style style:name="P" style:family="text"><style:text-properties style:text-position="{value}"/></style:style>"#
            );
            assert_eq!(
                read_fixture(Fixture {
                    styles: &styles,
                    body: r#"<text:p>a<text:span text:style-name="P">b</text:span></text:p>"#,
                    ..Fixture::default()
                })
                .unwrap_err(),
                unsupported("text-position"),
                "{value}"
            );
        }
    }

    #[test]
    fn rejects_spacing_scaling_resizing_and_justified_breaks() {
        let styles = r#"<style:style style:name="Zero" style:family="text"><style:text-properties style:text-position="0% 0%"/></style:style><style:style style:name="Inert" style:family="text"><style:text-properties style:text-position="0% 100%" fo:letter-spacing="0cm" style:text-scale="100%"/></style:style><style:style style:name="Tight" style:family="text"><style:text-properties fo:letter-spacing="-1pt"/></style:style><style:style style:name="Wide" style:family="text"><style:text-properties fo:letter-spacing="1pt"/></style:style><style:style style:name="Flat" style:family="text"><style:text-properties style:text-scale="0%"/></style:style><style:style style:name="Half" style:family="text"><style:text-properties style:text-scale="50%"/></style:style><style:style style:name="J" style:family="paragraph"><style:paragraph-properties fo:text-align="justify"/></style:style>"#;
        let read = |body: &str| {
            read_fixture(Fixture {
                styles,
                body,
                ..Fixture::default()
            })
        };
        let span = |style: &str| {
            format!(r#"<text:p>a<text:span text:style-name="{style}">b</text:span></text:p>"#)
        };
        assert_eq!(
            read(&span("Zero")).unwrap_err(),
            unsupported("text-position")
        );
        assert!(read(&span("Inert")).is_ok());
        for style in ["Tight", "Wide"] {
            assert_eq!(
                read(&span(style)).unwrap_err(),
                unsupported("letter-spacing"),
                "{style}"
            );
        }
        for style in ["Flat", "Half"] {
            assert_eq!(
                read(&span(style)).unwrap_err(),
                unsupported("text-scale"),
                "{style}"
            );
        }
        // Justified paragraphs map to DOCX `both`; LibreOffice stretches a
        // line ended by a manual break, the DOCX layout does not.
        let justified =
            read(r#"<text:p text:style-name="J">alpha beta gamma delta</text:p>"#).unwrap();
        assert_eq!(paragraphs(&justified)[0].alignment, "justify");
        assert_eq!(
            read(r#"<text:p text:style-name="J">alpha beta<text:line-break/>gamma delta</text:p>"#)
                .unwrap_err(),
            unsupported("justify-with-line-break")
        );
    }

    #[test]
    fn contextual_spacing_stays_within_a_content_area() {
        let styles = r#"<style:style style:name="C" style:family="paragraph"><style:paragraph-properties fo:margin-top="10pt" fo:margin-bottom="20pt" style:contextual-spacing="true"/></style:style>"#;
        let gaps = |body: &str| -> Vec<f64> {
            let document = read_fixture(Fixture {
                styles,
                body,
                ..Fixture::default()
            })
            .unwrap();
            paragraphs(&document)
                .iter()
                .map(|p| p.space_before)
                .collect()
        };
        let p = r#"<text:p text:style-name="C">x</text:p>"#;
        assert_eq!(gaps(&format!("{p}{p}")), [10.0, 0.0]);
        // Entering, leaving and nested sections are other content areas.
        assert_eq!(
            gaps(&format!(
                "{p}<text:section text:name=\"s\">{p}</text:section>"
            )),
            [10.0, 30.0]
        );
        assert_eq!(
            gaps(&format!(
                "<text:section text:name=\"s\">{p}</text:section>{p}"
            )),
            [10.0, 30.0]
        );
        assert_eq!(
            gaps(&format!("<text:section text:name=\"a\">{p}<text:section text:name=\"b\">{p}</text:section>{p}</text:section>")),
            [10.0, 30.0, 30.0]
        );
        assert_eq!(
            gaps(&format!(
                "<text:section text:name=\"s\">{p}{p}</text:section>"
            )),
            [10.0, 0.0]
        );
    }

    #[test]
    fn rejects_font_sizes_below_one_point() {
        for size in ["0pt", "-12pt", "0.01pt", "0.99pt"] {
            let styles = format!(
                r#"<style:style style:name="T" style:family="text"><style:text-properties fo:font-size="{size}"/></style:style>"#
            );
            assert_eq!(
                read_fixture(Fixture {
                    styles: &styles,
                    body: r#"<text:p>a<text:span text:style-name="T">b</text:span></text:p>"#,
                    ..Fixture::default()
                })
                .unwrap_err(),
                unsupported("font-size-range"),
                "{size}"
            );
        }
    }

    #[test]
    fn rejects_page_layouts_without_a_text_area_and_oversized_fonts() {
        for page in [
            r#" fo:page-width="-12pt""#,
            r#" fo:page-height="0cm""#,
            r#" fo:margin-left="-1cm""#,
            r#" fo:margin-left="10cm" fo:margin-right="12cm""#,
        ] {
            let error = read_fixture(Fixture {
                page,
                ..Fixture::default()
            })
            .unwrap_err();
            assert!(error.contains("no text area"), "{page}: {error}");
        }
        for (size, ok) in [("1000pt", true), ("1001pt", false), ("3000pt", false)] {
            let styles = format!(
                r#"<style:style style:name="T" style:family="text"><style:text-properties fo:font-size="{size}"/></style:style>"#
            );
            let result = read_fixture(Fixture {
                styles: &styles,
                body: r#"<text:p>a<text:span text:style-name="T">b</text:span></text:p>"#,
                ..Fixture::default()
            });
            assert_eq!(result.is_ok(), ok, "{size}");
        }
    }

    #[test]
    fn compatibility_rules_apply_within_their_measured_scope() {
        let on = LIBREOFFICE_RULES;
        // Asian and complex defaults differ from the Latin face, so a weak
        // character's class matters.
        let mixed = r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="12pt" style:font-name-complex="Sans" style:font-size-complex="12pt" style:letter-kerning="true""#;
        let read = |body: &str, styles: &str, rules: CompatRules| {
            read_fixture_with(
                Fixture {
                    default_text: mixed,
                    styles,
                    body,
                    ..Fixture::default()
                },
                rules,
            )
        };
        for body in [
            "<text:p>a b\u{2014}c\u{1f600}</text:p>",
            "<text:p>\u{201c}a\u{201d}</text:p>",
        ] {
            assert_eq!(
                read(body, "", SPEC_ONLY).unwrap_err(),
                unsupported("script-unmapped-character"),
                "{body}"
            );
            let document = read(body, "", on).unwrap();
            assert!(
                texts(paragraphs(&document)[0]).len() == 1,
                "weak characters take the Latin class: {body}"
            );
        }
        // Emoji-sequence components are weak; other unmapped marks and
        // format characters (zero-width space, a combining mark outside
        // Table 22) are not and keep rejecting.
        assert!(read("<text:p>a\u{2764}\u{fe0f}b</text:p>", "", on).is_ok());
        for body in ["<text:p>a\u{200b}b</text:p>", "<text:p>a\u{1ab0}b</text:p>"] {
            assert_eq!(
                read(body, "", on).unwrap_err(),
                unsupported("script-unmapped-character"),
                "{body}"
            );
        }
        // An unmapped letter is not weak.
        assert_eq!(
            read("<text:p>a\u{0500}</text:p>", "", on).unwrap_err(),
            unsupported("script-unmapped-character")
        );

        let style = |height: &str| {
            format!(
                r#"<style:style style:name="L" style:family="paragraph"><style:paragraph-properties fo:line-height="{height}"/></style:style><style:style style:name="Big" style:family="text"><style:text-properties fo:font-size="24pt"/></style:style>"#
            )
        };
        let p = r#"<text:p text:style-name="L">x y</text:p>"#;
        assert_eq!(
            read(p, &style("115%"), SPEC_ONLY).unwrap_err(),
            unsupported("line-height-proportional")
        );
        let document = read(p, &style("115%"), on).unwrap();
        let spacing = paragraphs(&document)[0].line_spacing.clone().unwrap();
        assert_eq!((spacing.value, spacing.rule.as_str()), (1.15, "auto"));
        for height in ["40%", "301%"] {
            assert_eq!(
                read(p, &style(height), on).unwrap_err(),
                unsupported("line-height-proportional")
            );
        }
        assert_eq!(
            read(r#"<text:p text:style-name="L">x <text:span text:style-name="Big">y</text:span></text:p>"#, &style("115%"), on)
                .unwrap_err(),
            unsupported("line-height-proportional-mixed")
        );
    }

    #[test]
    fn soft_hyphens_become_optional_hyphen_runs() {
        let styles = r#"<style:style style:name="R" style:family="text"><style:text-properties fo:font-size="20pt"/></style:style><style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>"#;
        assert_eq!(
            read_fixture(Fixture {
                styles,
                body: "<text:p>abc<text:span text:style-name=\"R\">\u{ad}</text:span>def</text:p>",
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("soft-hyphen-formatting")
        );
        let document = read_fixture(Fixture {
            styles,
            body: "<text:p>abc\u{ad}def<text:span text:style-name=\"B\">\u{ad}</text:span>gh</text:p>",
            ..Fixture::default()
        })
        .unwrap();
        let runs: Vec<(String, bool, f64)> = paragraphs(&document)[0]
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(t) => (t.text.clone(), t.optional_hyphen, t.font_size),
                _ => panic!("text runs only"),
            })
            .collect();
        assert_eq!(
            runs,
            [
                ("abc".into(), false, 12.0),
                (String::new(), true, 12.0),
                ("def".into(), false, 12.0),
                (String::new(), true, 12.0),
                ("gh".into(), false, 12.0),
            ]
        );
        for body in [
            "<text:p>abc\u{ad}<text:line-break/>def</text:p>",
            "<text:p>abc\u{ad}\u{ad}</text:p>",
        ] {
            assert!(
                read_fixture(Fixture {
                    body,
                    ..Fixture::default()
                })
                .is_ok(),
                "{body}"
            );
        }
    }

    #[test]
    fn font_color_follows_the_nearest_style_under_the_approved_rule() {
        let auto = r#"style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Serif" style:font-size-asian="12pt" style:font-name-complex="Serif" style:font-size-complex="12pt" style:letter-kerning="true" style:use-window-font-color="true" loext:opacity="0%""#;
        let styles = r##"<style:style style:name="R" style:family="text"><style:text-properties fo:color="#ff0000" loext:opacity="100%"/></style:style><style:style style:name="A" style:family="text" style:parent-style-name="R"><style:text-properties style:use-window-font-color="true"/></style:style><style:style style:name="Z" style:family="text"><style:text-properties fo:color="#ff0000" loext:opacity="0%"/></style:style><style:style style:name="H" style:family="text"><style:text-properties fo:color="#ff0000" loext:opacity="50%"/></style:style>"##;
        let read = |body: &str, rules: CompatRules| {
            read_fixture_with(
                Fixture {
                    default_text: auto,
                    styles,
                    body,
                    ..Fixture::default()
                },
                rules,
            )
        };
        let nearest = LIBREOFFICE_RULES;
        let body = r#"<text:p>a<text:span text:style-name="R">b</text:span><text:span text:style-name="A">c</text:span></text:p>"#;
        // An explicit span colour under an inherited automatic flag: the
        // readings differ, so it rejects without the rule.
        assert_eq!(
            read(body, SPEC_ONLY).unwrap_err(),
            unsupported("font-color-cascade")
        );
        let document = read(body, nearest).unwrap();
        let colors: Vec<(Option<String>, bool)> = paragraphs(&document)[0]
            .runs
            .iter()
            .map(|run| match run {
                DocRun::Text(t) => (t.color.clone(), t.color_auto),
                _ => panic!(),
            })
            .collect();
        assert_eq!(
            colors,
            [(None, true), (Some("ff0000".into()), false), (None, true)]
        );
        // A style setting both attributes rejects wherever it is applied,
        // also under a nearer style that decides the colour.
        let ambiguous = r##"<style:style style:name="P" style:family="text"><style:text-properties fo:color="#ff0000" style:use-window-font-color="true"/></style:style><style:style style:name="C" style:family="text" style:parent-style-name="P"><style:text-properties fo:color="#0000ff"/></style:style>"##;
        for style in ["P", "C"] {
            assert_eq!(
                read_fixture_with(
                    Fixture {
                        default_text: auto,
                        styles: ambiguous,
                        body: &format!(
                            r#"<text:p><text:span text:style-name="{style}">b</text:span></text:p>"#
                        ),
                        ..Fixture::default()
                    },
                    nearest,
                )
                .unwrap_err(),
                unsupported("font-color-ambiguous"),
                "{style}"
            );
        }
        // Automatic text alone agrees in both readings.
        assert!(read("<text:p>plain</text:p>", LIBREOFFICE_RULES).is_ok());
        // Opacity follows the resolved colour: 0% only with automatic colour.
        for style in ["Z", "H"] {
            assert_eq!(
                read(
                    &format!(
                        r#"<text:p><text:span text:style-name="{style}">b</text:span></text:p>"#
                    ),
                    nearest
                )
                .unwrap_err(),
                unsupported("text-opacity"),
                "{style}"
            );
        }
    }

    #[test]
    fn paragraph_backgrounds_with_libreoffice_area_fill() {
        let style = |name: &str, background: &str, fill: &str, margin: &str| {
            format!(
                r#"<style:style style:name="{name}" style:family="paragraph"><loext:graphic-properties {fill}/><style:paragraph-properties fo:background-color="{background}" fo:margin-bottom="{margin}"/></style:style>"#
            )
        };
        let read = |styles: &str, body: &str| {
            read_fixture(Fixture {
                styles,
                body,
                ..Fixture::default()
            })
        };
        let solid = r##"draw:fill="solid" draw:fill-color="#FFFFC8" draw:opacity="100%""##;
        // A transparent background with fill none, and a solid fill that
        // repeats the background, map to the paragraph shading.
        let styles = style("N", "transparent", r#"draw:fill="none""#, "0pt")
            + &style("Y", "#ffffc8", solid, "0pt");
        let document = read(&styles, r#"<text:p text:style-name="N">a</text:p><text:p text:style-name="Y">b</text:p><text:p text:style-name="Y">c</text:p>"#).unwrap();
        let shading: Vec<Option<String>> = paragraphs(&document)
            .iter()
            .map(|p| p.shading.clone())
            .collect();
        assert_eq!(
            shading,
            [None, Some("ffffc8".into()), Some("ffffc8".into())]
        );
        // Disagreement, other fill kinds and partial opacity reject.
        for (background, fill) in [
            (
                "#ffffc8",
                r##"draw:fill="solid" draw:fill-color="#c8ffff""##,
            ),
            ("#ffffc8", r#"draw:fill="none""#),
            (
                "transparent",
                r##"draw:fill="solid" draw:fill-color="#ffffc8""##,
            ),
            (
                "#ffffc8",
                r##"draw:fill="solid" draw:fill-color="#ffffc8" draw:opacity="50%""##,
            ),
            (
                "#ffffc8",
                r#"draw:fill="gradient" draw:fill-gradient-name="G""#,
            ),
        ] {
            assert_eq!(
                read(
                    &style("X", background, fill, "0pt"),
                    r#"<text:p text:style-name="X">x</text:p>"#
                )
                .unwrap_err(),
                unsupported("paragraph-fill"),
                "{background} {fill}"
            );
        }
        // LibreOffice fills the gap between consecutive filled paragraphs;
        // the DOCX layout does not, so a gap rejects.
        assert_eq!(
            read(
                &style("G", "#ffffc8", solid, "6pt"),
                r#"<text:p text:style-name="G">a</text:p><text:p text:style-name="G">b</text:p>"#
            )
            .unwrap_err(),
            unsupported("paragraph-fill-gap")
        );
        assert!(read(
            &style("G", "#ffffc8", solid, "6pt"),
            r#"<text:p text:style-name="G">a</text:p><text:p>b</text:p>"#
        )
        .is_ok());
        // Different colours, and a section boundary, leave the gap unfilled
        // in LibreOffice too.
        let two = style("G", "#ffffc8", solid, "6pt")
            + &style(
                "H",
                "#c8ffff",
                r##"draw:fill="solid" draw:fill-color="#c8ffff""##,
                "6pt",
            );
        assert!(read(
            &two,
            r#"<text:p text:style-name="G">a</text:p><text:p text:style-name="H">b</text:p>"#
        )
        .is_ok());
        assert!(read(&two, r#"<text:p text:style-name="G">a</text:p><text:section text:name="s"><text:p text:style-name="G">b</text:p></text:section>"#).is_ok());
        // A fully opaque percentage in any spelling repeats the background.
        assert!(read(
            &style(
                "O",
                "#ffffc8",
                r##"draw:fill="solid" draw:fill-color="#ffffc8" draw:opacity="100.0%""##,
                "0pt"
            ),
            r#"<text:p text:style-name="O">x</text:p>"#
        )
        .is_ok());
        // A default-style background fills the page in LibreOffice.
        assert_eq!(
            read_fixture(Fixture {
                default_paragraph: r##" fo:background-color="#ffffc8""##,
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("default-style-background")
        );
        // The default-style background rejects even when every paragraph
        // overrides it.
        assert_eq!(
            read_fixture(Fixture {
                default_paragraph: r##" fo:background-color="#ffffc8""##,
                styles: r#"<style:style style:name="T" style:family="paragraph"><style:paragraph-properties fo:background-color="transparent"/></style:style>"#,
                body: r#"<text:p text:style-name="T">x</text:p>"#,
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("default-style-background")
        );
        // Children of the area fill (e.g. a background image) reject.
        let image = r##"<style:style style:name="I" style:family="paragraph"><loext:graphic-properties draw:fill="none"><style:background-image/></loext:graphic-properties></style:style>"##;
        assert_eq!(
            read(image, r#"<text:p text:style-name="I">x</text:p>"#).unwrap_err(),
            unsupported("paragraph-fill")
        );
        // Automatic (and unset) text colour is resolved as LibreOffice paints
        // it: over its own run shading first, else the paragraph background.
        let colors = |document: &Document| -> Vec<(Option<String>, bool)> {
            paragraphs(document)
                .iter()
                .flat_map(|p| p.runs.iter())
                .filter_map(|run| match run {
                    DocRun::Text(text) => Some((text.color.clone(), text.color_auto)),
                    _ => None,
                })
                .collect()
        };
        let run = |shade: &str| {
            format!(
                r##"<style:style style:name="R" style:family="text"><style:text-properties fo:background-color="{shade}"/></style:style><style:style style:name="E" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style>"##
            ) + &style(
                "Y",
                "#ffffc8",
                r##"draw:fill="solid" draw:fill-color="#ffffc8""##,
                "0pt",
            )
        };
        let body = r#"<text:p text:style-name="Y">a<text:span text:style-name="R">b</text:span><text:span text:style-name="E">c</text:span></text:p>"#;
        assert_eq!(
            colors(&read(&run("#999999"), body).unwrap()),
            [
                (Some("000000".into()), false),
                (Some("ffffff".into()), false),
                // An explicit colour is kept.
                (Some("ff0000".into()), false),
            ]
        );
        // LibreOffice's IsDark(): truncated WCAG luminance x 255 <= 87, and
        // the integer luminance <= 62 for #729fcf only.
        for (background, painted) in [
            ("#9e9e9e", "ffffff"),
            ("#9f9f9f", "000000"),
            ("#00c000", "000000"),
            ("#b0b000", "000000"),
            ("#8080ff", "ffffff"),
            ("#b8b8b8", "000000"),
            ("#ff6d6d", "ffffff"),
            ("#729fcf", "000000"),
            ("#729fce", "ffffff"),
            ("#000000", "ffffff"),
            ("#ffffff", "000000"),
        ] {
            let document = read(
                &style("A", background, "", "0pt"),
                r#"<text:p text:style-name="A">x</text:p>"#,
            )
            .unwrap();
            assert_eq!(
                colors(&document),
                [(Some(painted.to_string()), false)],
                "{background}"
            );
        }
        // Without a background automatic colour stays automatic.
        assert_eq!(
            colors(&read("", "<text:p>x</text:p>").unwrap()),
            [(None, false)]
        );
    }

    #[test]
    fn admits_libreoffice_layout_settings_inert_for_admitted_content() {
        let settings = |items: &[(&str, &str)]| {
            format!(
                r#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings">{}</config:config-item-set></office:settings></office:document-settings>"#,
                items
                    .iter()
                    .map(|(k, v)| format!(r#"<config:config-item config:name="{k}" config:type="boolean">{v}</config:config-item>"#))
                    .collect::<String>()
            )
        };
        // The flags Word-derived LibreOffice documents carry, in either value.
        for value in ["true", "false"] {
            let items: Vec<(&str, &str)> =
                INERT_SETTINGS.iter().map(|name| (*name, value)).collect();
            assert!(
                check_settings(&settings(&items), Some("1.3")).is_ok(),
                "{value}"
            );
        }
        // Pinned items and non-boolean values still reject.
        for items in [
            [("PrinterIndependentLayout", "disabled")],
            [("MsWordCompTrailingBlanks", "maybe")],
        ] {
            assert_eq!(
                check_settings(&settings(&items), Some("1.3")).unwrap_err(),
                unsupported("compatibility-setting")
            );
        }
        // Boolean items must declare config:type="boolean" (§19.30).
        for item in [
            r#"<config:config-item config:name="MsWordCompTrailingBlanks" config:type="string">true</config:config-item>"#,
            r#"<config:config-item config:name="MsWordCompTrailingBlanks">true</config:config-item>"#,
            r#"<config:config-item config:name="ParaSpaceMax" config:type="string">false</config:config-item>"#,
        ] {
            let text = settings(&[]).replace(
                "</config:config-item-set>",
                &format!("{item}</config:config-item-set>"),
            );
            assert_eq!(
                check_settings(&text, Some("1.3")).unwrap_err(),
                unsupported("compatibility-setting"),
                "{item}"
            );
        }
        for value in ["true", "false"] {
            assert_eq!(
                check_settings(
                    &settings(&[("BalanceSpacesAndIdeographicSpaces", value)]),
                    Some("1.3")
                )
                .unwrap()
                .balance,
                value == "true"
            );
        }
        // Justification settings: shrinking rejects justified paragraphs;
        // DoNotJustifyLinesWithManualBreak admits their manual breaks.
        let justified = |settings: &str, body: &str| {
            let settings = format!(
                r#"<config:config-item config:name="{settings}" config:type="boolean">true</config:config-item>"#
            );
            read_fixture(Fixture {
                styles: r#"<style:style style:name="J" style:family="paragraph"><style:paragraph-properties fo:text-align="justify"/></style:style>"#,
                body,
                settings: &settings,
                ..Fixture::default()
            })
        };
        assert_eq!(
            justified(
                "JustifyLinesWithShrinking",
                r#"<text:p text:style-name="J">alpha beta</text:p>"#
            )
            .unwrap_err(),
            unsupported("justify-shrinking")
        );
        assert!(justified(
            "DoNotJustifyLinesWithManualBreak",
            r#"<text:p text:style-name="J">alpha<text:line-break/>beta</text:p>"#
        )
        .is_ok());
        // Left-aligned paragraphs are unaffected by either setting.
        assert!(justified("JustifyLinesWithShrinking", "<text:p>alpha beta</text:p>").is_ok());
    }

    #[test]
    fn balanced_spaces_follow_libreoffice_portions() {
        let styles = r##"<style:style style:name="Red" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="Empty" style:family="text"/><style:style style:name="J" style:family="paragraph"><style:paragraph-properties fo:text-align="justify"/></style:style><style:style style:name="Red2" style:family="text" style:display-name="Red"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="PRed" style:family="paragraph"><style:text-properties fo:color="#ff0000"/></style:style>"##;
        let automatic = r##"<style:style style:name="T1" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="T2" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="TE" style:family="text"><style:text-properties/></style:style><style:style style:name="R1" style:family="text"><style:text-properties officeooo:rsid="00111111"/></style:style><style:style style:name="R2" style:family="text"><style:text-properties officeooo:rsid="00222222"/></style:style><style:style style:name="NR" style:family="text" style:parent-style-name="Red"><style:text-properties officeooo:rsid="00333333"/></style:style><style:style style:name="De" style:family="text"><style:text-properties fo:language="de" fo:country="DE"/></style:style><style:style style:name="A10" style:family="text"><style:text-properties style:font-size-asian="10pt"/></style:style><style:style style:name="En" style:family="text"><style:text-properties fo:language="en"/></style:style><style:style style:name="Zh" style:family="text"><style:text-properties style:language-asian="zh" style:country-asian="CN"/></style:style><style:style style:name="Cmn" style:family="text"><style:text-properties style:language-asian="cmn" style:country-asian="CN"/></style:style><style:style style:name="Us" style:family="text"><style:text-properties fo:country="US"/></style:style><style:style style:name="Gb" style:family="text"><style:text-properties fo:country="GB"/></style:style><style:style style:name="EnNone" style:family="text"><style:text-properties fo:language="en" fo:country="none"/></style:style><style:style style:name="A12x" style:family="text"><style:text-properties style:font-size-asian="12.024pt"/></style:style><style:style style:name="TU" style:family="text"><style:text-properties fo:color="#FF0000"/></style:style><style:style style:name="TR" style:family="text" style:parent-style-name="Red"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="PA" style:family="paragraph"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="S12" style:family="text"><style:text-properties fo:font-size="12pt"/></style:style><style:style style:name="S12x" style:family="text"><style:text-properties fo:font-size="12.001pt"/></style:style>"##;
        let read = |balance: bool, body: &str| {
            read_fixture(Fixture {
                styles,
                automatic,
                body,
                settings: &format!(
                    r#"<config:config-item config:name="BalanceSpacesAndIdeographicSpaces" config:type="boolean">{balance}</config:config-item>"#
                ),
                ..Fixture::default()
            })
        };
        // Each U+0020's advance, in order (Asian size 12 pt: 6 pt).
        let advances = |balance: bool, body: &str| -> Vec<Option<f64>> {
            let document = read(balance, body).unwrap_or_else(|e| panic!("{body}: {e}"));
            paragraphs(&document)[0]
                .runs
                .iter()
                .flat_map(|run| match run {
                    DocRun::Text(text) => text
                        .text
                        .chars()
                        .filter(|c| *c == ' ')
                        .map(|_| text.space_advance_pt)
                        .collect::<Vec<_>>(),
                    _ => Vec::new(),
                })
                .collect()
        };
        let s = |style: &str, text: &str| {
            format!(r#"<text:span text:style-name="{style}">{text}</text:span>"#)
        };
        let six = Some(6.0);
        for (body, expected) in [
            // An inner space in one portion keeps its glyph advance.
            ("aa bb cc".to_string(), vec![None, None]),
            // Consecutive spaces, and a space at either edge of a portion.
            ("aa<text:s text:c=\"2\"/>bb".into(), vec![six, six]),
            (format!("aa{}bb cc", s("Red", " ")), vec![six, None]),
            (format!("aa {}", s("Red", "bb")), vec![six]),
            (format!("{} bb cc", s("Red", "aa")), vec![six, None]),
            // A named style without properties is a portion; an automatic
            // style without properties, or with revision ids only, is not.
            (format!("aa{}bb", s("Empty", " ")), vec![six]),
            (format!("aa{}bb", s("TE", " ")), vec![None]),
            (format!("{}{} cc", s("R1", "aa "), s("R2", "bb")), vec![None, None]),
            // Equal adjacent formatting is one portion; a character style is
            // not equal to the same properties set directly.
            (format!("{}{} cc", s("Red", "aa"), s("Red", " bb")), vec![None, six]),
            (format!("{}{} cc", s("T1", "aa"), s("T2", " bb")), vec![None, six]),
            (format!("{}{} cc", s("Red", "aa"), s("NR", " bb")), vec![None, six]),
            (format!("{}{} cc", s("Red", "aa"), s("T1", " bb")), vec![six, six]),
            // A repeated nested character style is a portion of its own.
            (s("Red", &format!("aa{} cc", s("Red", " bb"))), vec![six, six]),
            // Bookmarks and reference marks end portions; soft page breaks
            // do not. A kept collapsed space precedes the bookmark.
            ("aa<text:bookmark text:name=\"b\"/> bb".into(), vec![six]),
            ("aa <text:bookmark text:name=\"b\"/> bb".into(), vec![six]),
            ("aa<text:reference-mark-start text:name=\"r\"/> bb<text:reference-mark-end text:name=\"r\"/> cc".into(), vec![six, six]),
            ("aa<text:soft-page-break/> bb".into(), vec![None]),
            // Soft hyphens and line breaks end portions.
            ("aa \u{ad}bb cc".into(), vec![six, None]),
            ("aa<text:line-break/><text:s/>bb cc".into(), vec![six, None]),
            // The space's own Asian size.
            (format!("aa{}bb", s("A10", " ")), vec![Some(5.0)]),
            // No-break spaces are not balanced, but they, non-breaking
            // hyphens, zero-width spaces and word joiners are portions of
            // their own; other punctuation and spaces are not.
            ("aa\u{a0}bb".into(), vec![]),
            ("aa\u{a0} bb".into(), vec![six]),
            ("aa \u{2011}bb".into(), vec![six]),
            ("aa\u{200b} bb".into(), vec![six]),
            ("aa\u{2060} bb".into(), vec![six]),
            ("aa\u{202f} bb".into(), vec![None]),
            ("aa- bb".into(), vec![None]),
            // Styles are compared by display name.
            (format!("{}{} cc", s("Red", "aa"), s("Red2", " bb")), vec![None, six]),
        ] {
            assert_eq!(advances(true, &format!("<text:p>{body}</text:p>")), expected, "{body}");
            assert!(
                advances(false, &format!("<text:p>{body}</text:p>")).iter().all(Option::is_none),
                "{body}"
            );
        }
        // Seams LibreOffice decides by comparing formatting items after
        // merging (a property repeated by a nested span, spelled differently,
        // or repeating the paragraph's automatic formatting) reject next to
        // a space that would otherwise keep its advance.
        for body in [
            s("T1", &format!("aa{} cc", s("T2", " bb"))),
            format!("{}{}", s("T1", "aa"), s("TU", " bb")),
            format!(
                r#"<text:p text:style-name="PA">aa{}</text:p>"#,
                s("T1", " bb")
            ),
            // Sizes LibreOffice may round together.
            format!("{}{}", s("S12", "aa"), s("S12x", " bb")),
            // An empty formatted span, a byte order mark, a direction mark.
            format!("aa{} bb", s("Red", "")),
            // The same language spelled differently (a country of "none" is
            // no country): not a boundary in itself, and unknown otherwise.
            format!("{}{} cc", s("En", "aa"), s("EnNone", " bb")),
            // A different language (LibreOffice maps some locales together).
            s("En", &format!("aa{}bb", s("De", " "))),
            // Asian or complex languages.
            format!("{}{} cc", s("Zh", "aa"), s("Cmn", " bb")),
            // A country without a language.
            format!("{}{} cc", s("Us", "aa"), s("Gb", " bb")),
            // A formatted space merged into a collapsed space.
            format!("aa {}bb", s("Red", " ")),
            "aa\u{feff} bb".into(),
            "aa \u{200e}bb".into(),
        ] {
            let body = if body.starts_with("<text:p") {
                body
            } else {
                format!("<text:p>{body}</text:p>")
            };
            assert_eq!(
                read(true, &body).unwrap_err(),
                unsupported("balanced-space-portion"),
                "{body}"
            );
            assert_ne!(
                read(false, &body).err(),
                Some(unsupported("balanced-space-portion")),
                "{body}"
            );
        }
        // Asian sizes between LibreOffice's twips reject.
        assert_eq!(
            read(true, &format!("<text:p>aa{}bb</text:p>", s("A12x", " "))).unwrap_err(),
            unsupported("balanced-space-size")
        );
        // A character style nested in an automatic style is not the same as
        // an automatic style with that parent and the same property.
        assert_eq!(
            read(
                true,
                &format!(
                    "<text:p>{}{}</text:p>",
                    s("T1", &s("Red", "aa")),
                    s("TR", " bb")
                )
            )
            .unwrap_err(),
            unsupported("balanced-space-portion")
        );
        // Paragraph formatting does not take part otherwise: a span repeating
        // a named paragraph style's formatting is a portion of its own.
        assert_eq!(
            advances(
                true,
                &format!(
                    r#"<text:p text:style-name="PRed">aa{}bb</text:p>"#,
                    s("Red", " ")
                )
            ),
            vec![six]
        );
        // A balanced space in a justified paragraph rejects unless it hangs at
        // the paragraph's end.
        assert_eq!(
            read(
                true,
                r#"<text:p text:style-name="J">aa<text:s text:c="2"/>bb</text:p>"#
            )
            .unwrap_err(),
            unsupported("justify-balanced-spaces")
        );
        assert!(read(
            true,
            r#"<text:p text:style-name="J">aa bb<text:s text:c="2"/></text:p>"#
        )
        .is_ok());
        assert!(read(
            false,
            r#"<text:p text:style-name="J">aa<text:s text:c="2"/>bb</text:p>"#
        )
        .is_ok());
    }

    #[test]
    fn spaces_size_lines_as_in_libreoffice() {
        let styles = r#"<style:style style:name="Big" style:family="text"><style:text-properties fo:font-size="30pt"/></style:style><style:style style:name="Mid" style:family="text"><style:text-properties fo:font-size="20pt"/></style:style><style:style style:name="Small" style:family="text"><style:text-properties fo:font-size="10pt"/></style:style><style:style style:name="Bold" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="BoldBig" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-size="30pt"/></style:style>"#;
        let item = |name: &str, value: bool| {
            format!(
                r#"<config:config-item config:name="{name}" config:type="boolean">{value}</config:config-item>"#
            )
        };
        let read = |ignore: bool, blanks: bool, body: &str| {
            let settings = item("IgnoreTabsAndBlanksForLineCalculation", ignore)
                + &item("MsWordCompTrailingBlanks", blanks);
            read_fixture(Fixture {
                styles,
                body,
                settings: &settings,
                ..Fixture::default()
            })
        };
        // A space that is larger than, or in another face than, the preceding
        // text on its line, in each position.
        let mid = [
            r#"<text:p>a<text:span text:style-name="Big"> </text:span>b</text:p>"#,
            // Every white space character, not only U+0020 and U+00A0.
            "<text:p>a<text:span text:style-name=\"Big\">\u{2003}</text:span>b</text:p>",
            // Weight and posture are part of the face.
            r#"<text:p><text:span text:style-name="BoldBig">A</text:span><text:span text:style-name="Mid"> </text:span>b</text:p>"#,
            // The comparison text must be on the space's own line.
            r#"<text:p><text:span text:style-name="Big">A<text:line-break/></text:span><text:span text:style-name="Mid"><text:s/></text:span>b</text:p>"#,
            r#"<text:p><text:span text:style-name="Big"><text:s/></text:span>a</text:p>"#,
        ];
        let trailing = [
            r#"<text:p>a b<text:span text:style-name="Big"><text:s text:c="3"/></text:span></text:p>"#,
            r#"<text:p>a<text:span text:style-name="Bold"><text:s/></text:span></text:p>"#,
        ];
        // The spec-removed trailing space, which LibreOffice keeps.
        let removed = [r#"<text:p>a b<text:span text:style-name="Big"> </text:span></text:p>"#];
        let fitting = [
            r#"<text:p>a<text:span text:style-name="Small"> </text:span>b<text:span text:style-name="Small"><text:s/></text:span></text:p>"#,
            r#"<text:p><text:span text:style-name="Bold">a </text:span>b </text:p>"#,
        ];
        // (IgnoreTabsAndBlanks, TrailingBlanks) → which cases reject: where
        // LibreOffice does not size a line by a space the DOCX layout counts.
        for (ignore, blanks, mid_rejects, trailing_rejects, removed_rejects) in [
            (false, false, false, false, true),
            (false, true, false, true, false),
            (true, false, true, true, false),
            (true, true, true, true, false),
        ] {
            for (bodies, rejects) in [
                (mid.as_slice(), mid_rejects),
                (trailing.as_slice(), trailing_rejects),
                (removed.as_slice(), removed_rejects),
                (fitting.as_slice(), false),
            ] {
                for body in bodies {
                    let result = read(ignore, blanks, body);
                    if rejects {
                        assert_eq!(
                            result.unwrap_err(),
                            unsupported("space-formatting"),
                            "{ignore} {blanks} {body}"
                        );
                    } else {
                        assert!(result.is_ok(), "{ignore} {blanks} {body} {result:?}");
                    }
                }
            }
            // A line of only white space rejects whatever the settings.
            assert_eq!(
                read(ignore, blanks, r#"<text:p><text:s text:c="2"/></text:p>"#).unwrap_err(),
                unsupported("space-formatting")
            );
        }
    }

    #[test]
    fn trailing_space_decorations_follow_the_trailing_blank_settings() {
        let styles = r##"<style:style style:name="H" style:family="text"><style:text-properties fo:background-color="#ffff00"/></style:style><style:style style:name="U" style:family="text"><style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style><style:style style:name="S" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>"##;
        let item = |name: &str, value: bool| {
            format!(
                r#"<config:config-item config:name="{name}" config:type="boolean">{value}</config:config-item>"#
            )
        };
        // (MsWordCompTrailingBlanks, MsWordUlTrailSpace) → admitted styles of
        // a space ending the paragraph or preceding a break.
        for (blanks, underline, admitted) in [
            (false, false, [].as_slice()),
            (false, true, [].as_slice()),
            (true, false, ["H", "U", "S"].as_slice()),
            (true, true, ["H"].as_slice()),
        ] {
            let settings =
                item("MsWordCompTrailingBlanks", blanks) + &item("MsWordUlTrailSpace", underline);
            for style in ["H", "U", "S"] {
                for body in [
                    format!(
                        r#"<text:p><text:span text:style-name="{style}">a<text:s/></text:span></text:p>"#
                    ),
                    format!(
                        r#"<text:p><text:span text:style-name="{style}">a<text:s text:c="2"/></text:span><text:line-break/>b</text:p>"#
                    ),
                    // §6.1.2 removes this collapsed space; LibreOffice keeps it.
                    format!(
                        r#"<text:p><text:span text:style-name="{style}">a </text:span></text:p>"#
                    ),
                    // The U+0020 suffix after a no-break space.
                    format!(
                        r#"<text:p><text:span text:style-name="{style}">a&#160;<text:s/></text:span></text:p>"#
                    ),
                ] {
                    let result = read_fixture(Fixture {
                        styles,
                        body: &body,
                        settings: &settings,
                        ..Fixture::default()
                    });
                    if admitted.contains(&style) {
                        assert!(result.is_ok(), "{blanks} {underline} {body} {result:?}");
                    } else {
                        assert_eq!(
                            result.unwrap_err(),
                            unsupported("trailing-space-decoration"),
                            "{blanks} {underline} {body}"
                        );
                    }
                }
                // Both decorate a trailing no-break or other Unicode space.
                for space in ["&#160;", "&#8194;", "&#8201;"] {
                    for body in [
                        format!(
                            r#"<text:p><text:span text:style-name="{style}">a{space}{space}</text:span></text:p>"#
                        ),
                        format!(
                            r#"<text:p><text:span text:style-name="{style}">a{space}</text:span><text:line-break/>b</text:p>"#
                        ),
                    ] {
                        let result = read_fixture(Fixture {
                            styles,
                            body: &body,
                            settings: &settings,
                            ..Fixture::default()
                        });
                        assert!(result.is_ok(), "{blanks} {underline} {body} {result:?}");
                    }
                }
                // Spaces between visible text are not trailing.
                let body = format!(
                    r#"<text:p><text:span text:style-name="{style}">a b</text:span><text:line-break/>c</text:p>"#
                );
                assert!(read_fixture(Fixture {
                    styles,
                    body: &body,
                    settings: &settings,
                    ..Fixture::default()
                })
                .is_ok());
            }
        }
        // Without settings.xml LibreOffice's defaults (false) apply.
        let body = r#"<text:p><text:span text:style-name="H">a<text:s/></text:span></text:p>"#;
        assert_eq!(
            read_fixture(Fixture {
                styles,
                body,
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("trailing-space-decoration")
        );
        // Only the top-level ooo:configuration-settings set counts: an item in
        // another set does not override it, and repeated sets or items reject.
        let blanks = item("MsWordCompTrailingBlanks", false);
        let shadow = format!(
            r#"{blanks}<config:config-item-set config:name="other">{}</config:config-item-set>"#,
            item("MsWordCompTrailingBlanks", true)
        );
        assert_eq!(
            read_fixture(Fixture {
                styles,
                body,
                settings: &shadow,
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("trailing-space-decoration")
        );
        for settings in [
            format!("{blanks}{}", item("MsWordCompTrailingBlanks", false)),
            format!(
                r#"{blanks}</config:config-item-set><config:config-item-set config:name="ooo:configuration-settings">"#
            ),
        ] {
            assert_eq!(
                read_fixture(Fixture {
                    styles,
                    body,
                    settings: &settings,
                    ..Fixture::default()
                })
                .unwrap_err(),
                unsupported("compatibility-setting")
            );
        }
    }

    #[test]
    fn rejects_repeated_element_valued_properties() {
        let tab_stops = "<style:tab-stops/>".repeat(2);
        let styles = format!(
            r#"<style:style style:name="T" style:family="paragraph"><style:paragraph-properties>{tab_stops}</style:paragraph-properties></style:style>"#
        );
        assert!(read_fixture(Fixture {
            styles: &styles,
            ..Fixture::default()
        })
        .unwrap_err()
        .contains("repeats an element-valued property"));
    }

    #[test]
    fn rejects_padding_and_outline_level_indents() {
        let reject = |fixture: Fixture<'_>| read_fixture(fixture).unwrap_err();
        assert_eq!(
            reject(Fixture {
                styles: r#"<style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:padding-left="5pt"/></style:style>"#,
                body: r#"<text:p text:style-name="P">x</text:p>"#,
                ..Fixture::default()
            }),
            unsupported("paragraph-property")
        );
        assert!(read_fixture(Fixture {
            styles: r#"<style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:padding="0cm" fo:border="none"/></style:style>"#,
            body: r#"<text:p text:style-name="P">x</text:p>"#,
            ..Fixture::default()
        })
        .is_ok());
        let outline = |level: &str| {
            format!(
                r#"<text:outline-style style:name="Outline"><text:outline-level-style text:level="1" style:num-format="">{level}</text:outline-level-style></text:outline-style>"#
            )
        };
        let admitted = outline(
            r#"<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="nothing"/></style:list-level-properties>"#,
        );
        assert!(read_fixture(Fixture {
            styles: &admitted,
            body: r#"<text:h text:outline-level="1">x</text:h>"#,
            ..Fixture::default()
        })
        .is_ok());
        for level in [
            r#"<style:list-level-properties text:space-before="1cm"/>"#,
            r#"<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="nothing" fo:margin-left="1cm"/></style:list-level-properties>"#,
            r#"<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties>"#,
            r#"<style:text-properties fo:font-size="20pt"/>"#,
        ] {
            assert_eq!(
                reject(Fixture {
                    styles: &outline(level),
                    body: r#"<text:h text:outline-level="1">x</text:h>"#,
                    ..Fixture::default()
                }),
                unsupported("outline-style"),
                "{level}"
            );
        }
    }

    #[test]
    fn package_parts_carry_the_document_version() {
        let settings = |version: &str| {
            format!(
                r#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="{version}"/>"#
            )
        };
        assert!(check_settings(&settings("1.3"), Some("1.3")).is_ok());
        assert_eq!(
            check_settings(&settings("1.2"), Some("1.3")).unwrap_err(),
            unsupported("odf-version")
        );
        assert!(check_settings(
            r#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>"#,
            Some("1.3")
        )
        .is_err());
    }

    /// A list level in label alignment mode (M, I and T in cm).
    fn list_level(level: u32, format: &str, extra: &str, m: f64, i: f64, t: f64) -> String {
        format!(
            r#"<text:list-level-style-number text:level="{level}" style:num-format="{format}" style:num-suffix="."{extra}><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="{t}cm" fo:text-indent="{i}cm" fo:margin-left="{m}cm"/></style:list-level-properties></text:list-level-style-number>"#
        )
    }

    fn list_styles(levels: &str) -> String {
        format!(r#"<text:list-style style:name="L">{levels}</text:list-style>"#)
    }

    const LIST_SETTINGS: &str = r#"<config:config-item config:name="TabsRelativeToIndent" config:type="boolean">true</config:config-item>"#;

    fn read_list(styles: &str, body: &str) -> Result<Document, String> {
        read_fixture(Fixture {
            styles,
            body,
            settings: LIST_SETTINGS,
            ..Fixture::default()
        })
    }

    /// Each paragraph's label text (`None` for an unlabelled paragraph).
    fn labels(document: &Document) -> Vec<Option<String>> {
        paragraphs(document)
            .iter()
            .map(|paragraph| match paragraph.runs.first() {
                Some(DocRun::Text(text)) if text.list_label.is_some() => Some(text.text.clone()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn numbers_lists_as_libreoffice() {
        // Three levels starting at 5, 3 and 1, each showing all its levels.
        let styles = list_styles(
            &(list_level(1, "1", r#" text:start-value="5""#, 2.0, -1.0, 2.0)
                + &list_level(
                    2,
                    "1",
                    r#" text:start-value="3" text:display-levels="2""#,
                    3.0,
                    -1.0,
                    3.0,
                )
                + &list_level(3, "1", r#" text:display-levels="3""#, 4.0, -1.0, 4.0)),
        );
        let list = |attributes: &str, items: &str| {
            format!(r#"<text:list text:style-name="L"{attributes}>{items}</text:list>"#)
        };
        let item = |content: &str| format!("<text:list-item>{content}</text:list-item>");
        let p = |text: &str| format!("<text:p>{text}</text:p>");
        let nested = |items: &str| format!("<text:list>{items}</text:list>");
        let wrap = |content: &str| item(&nested(content));
        let cases = [
            // Items count on; a later item paragraph is unlabelled.
            (
                list("", &(item(&(p("a") + &p("a more"))) + &item(&p("b")))),
                vec![Some("5."), None, Some("6.")],
            ),
            // Word heading numbering: one list per heading, continued, with
            // wrapper items for the deeper levels (measured 5, 6, 6.3, 6.3.1,
            // 6.4, 7, 7.3.1).
            (
                list("", &item(&p("h1")))
                    + &list(r#" text:continue-numbering="true""#, &item(&p("h2")))
                    + &list(
                        r#" text:continue-numbering="true""#,
                        &wrap(&item(&p("h21"))),
                    )
                    + &list(
                        r#" text:continue-numbering="true""#,
                        &wrap(&item(&nested(&item(&p("h211"))))),
                    )
                    + &list(
                        r#" text:continue-numbering="true""#,
                        &wrap(&item(&p("h22"))),
                    )
                    + &list(r#" text:continue-numbering="true""#, &item(&p("h3")))
                    + &list(
                        r#" text:continue-numbering="true""#,
                        &wrap(&item(&nested(&item(&p("h311"))))),
                    ),
                vec![
                    Some("5."),
                    Some("6."),
                    Some("6.3."),
                    Some("6.3.1."),
                    Some("6.4."),
                    Some("7."),
                    Some("7.3.1."),
                ],
            ),
            // A wrapper's start value is ignored; its phantom takes the start.
            (
                list(
                    "",
                    r#"<text:list-item text:start-value="9"><text:list><text:list-item><text:p>w</text:p></text:list-item></text:list></text:list-item>"#,
                ),
                vec![Some("5.3.")],
            ),
            // Empty items and headers are not counted; a header's sublist
            // numbers through a phantom parent.
            (
                list(
                    "",
                    &(item("")
                        + "<text:list-header><text:p>h</text:p></text:list-header>"
                        + &item(&p("x"))),
                ),
                vec![None, Some("5.")],
            ),
            (
                list(
                    "",
                    &("<text:list-header>".to_string()
                        + &nested(&item(&p("e1")))
                        + "</text:list-header>"
                        + &item(&p("e2"))),
                ),
                vec![Some("5.3."), Some("6.")],
            ),
            // An item start sets the value; continuation by xml:id shares
            // the counters across another list.
            (
                list(
                    r#" xml:id="A""#,
                    r#"<text:list-item text:start-value="7"><text:p>s</text:p></text:list-item>"#,
                ) + &list("", &item(&p("other")))
                    + &list(r#" text:continue-list="A""#, &item(&p("t"))),
                vec![Some("7."), Some("5."), Some("8.")],
            ),
        ];
        for (body, expected) in cases {
            let document = read_list(&styles, &body).unwrap_or_else(|e| panic!("{body}: {e}"));
            let expected: Vec<Option<String>> = expected
                .into_iter()
                .map(|l| l.map(str::to_string))
                .collect();
            assert_eq!(labels(&document), expected, "{body}");
        }
    }

    #[test]
    fn composes_labels_as_libreoffice() {
        let one = |level1: &str, level2: &str, count: u32| -> String {
            let styles =
                format!(r#"<text:list-style style:name="L">{level1}{level2}</text:list-style>"#);
            let items = (0..count)
                .map(|_| "<text:list-item><text:p>x</text:p></text:list-item>")
                .collect::<String>();
            let body = format!(
                r#"<text:list text:style-name="L"><text:list-item><text:p>p</text:p><text:list>{items}</text:list></text:list-item></text:list>"#
            );
            let document = read_list(&styles, &body).unwrap_or_else(|e| panic!("{styles}: {e}"));
            labels(&document).last().cloned().flatten().unwrap()
        };
        let display2 =
            |format: &str| list_level(2, format, r#" text:display-levels="2""#, 3.0, -1.0, 3.0);
        let bullet1 = r#"<text:list-level-style-bullet text:level="1" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2cm" fo:text-indent="-1cm" fo:margin-left="2cm"/></style:list-level-properties></text:list-level-style-bullet>"#;
        // An empty-format ancestor is left out; a bullet ancestor leaves an
        // empty component (measured "1." and ".1.").
        assert_eq!(
            one(&list_level(1, "", "", 2.0, -1.0, 2.0), &display2("1"), 1),
            "1."
        );
        assert_eq!(one(bullet1, &display2("1"), 1), ".1.");
        // An empty own format shows only the affixes.
        assert_eq!(
            one(&list_level(1, "1", "", 2.0, -1.0, 2.0), &display2(""), 1),
            "."
        );
        // Formats.
        assert_eq!(
            one(&list_level(1, "1", "", 2.0, -1.0, 2.0), &display2("i"), 4),
            "1.iv."
        );
        assert_eq!(
            one(&list_level(1, "I", "", 2.0, -1.0, 2.0), &display2("A"), 3),
            "I.C."
        );
        // Zero renders "0" in every format, then counting goes on.
        let zero = |format: &str| list_level(2, format, r#" text:start-value="0""#, 3.0, -1.0, 3.0);
        assert_eq!(
            one(&list_level(1, "1", "", 2.0, -1.0, 2.0), &zero("a"), 1),
            "0."
        );
        assert_eq!(
            one(&list_level(1, "1", "", 2.0, -1.0, 2.0), &zero("i"), 2),
            "i."
        );
        // LibreOffice's label template must be that composition.
        let templated = |template: &str| {
            list_level(
                2,
                "1",
                &format!(r#" loext:num-list-format="{template}""#),
                3.0,
                -1.0,
                3.0,
            )
        };
        assert_eq!(
            one(
                &list_level(1, "1", "", 2.0, -1.0, 2.0),
                &templated("%1%.%2%."),
                1
            ),
            "1.1."
        );
        let styles =
            list_styles(&(list_level(1, "1", "", 2.0, -1.0, 2.0) + &templated("%2%-%1%.")));
        assert_eq!(
            read_list(&styles, r#"<text:list text:style-name="L"><text:list-item><text:p>p</text:p><text:list><text:list-item><text:p>x</text:p></text:list-item></text:list></text:list-item></text:list>"#).unwrap_err(),
            unsupported("list-format-template")
        );
    }

    #[test]
    fn takes_indents_from_linked_paragraph_styles() {
        let cm = 72.0 / 2.54;
        let styles = list_styles(&list_level(1, "1", "", 2.0, -1.0, 2.0))
            + r#"<style:style style:name="Base" style:family="paragraph" style:list-style-name="L"/><style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base"><style:paragraph-properties fo:margin-left="6cm" fo:text-indent="-2cm"/></style:style><style:style style:name="Plain" style:family="paragraph"><style:paragraph-properties fo:margin-left="6cm" fo:text-indent="-2cm"/></style:style>"#;
        let indents = |paragraph_style: &str| {
            let body = format!(
                r#"<text:list text:style-name="L"><text:list-item><text:p text:style-name="{paragraph_style}">a</text:p></text:list-item></text:list>"#
            );
            let document = read_list(&styles, &body).unwrap();
            let paragraph = paragraphs(&document)[0];
            (
                (paragraph.indent_left / cm * 100.0).round() / 100.0,
                (paragraph.indent_first / cm * 100.0).round() / 100.0,
            )
        };
        // A link inherited from the parent style applies the style's indents;
        // a style without one is ignored (measured).
        assert_eq!(indents("Child"), (6.0, -2.0));
        assert_eq!(indents("Plain"), (2.0, -1.0));
    }

    #[test]
    fn resolves_label_styles_in_the_list_style_scope() {
        // A common list style names a common character style even where a
        // content automatic text style shares its name.
        let styles = list_styles(&list_level(
            1,
            "1",
            r#" text:style-name="C""#,
            2.0,
            -1.0,
            2.0,
        )) + r#"<style:style style:name="C" style:family="text"><style:text-properties fo:font-size="20pt"/></style:style>"#;
        let document = read_fixture(Fixture {
            styles: &styles,
            automatic: r#"<style:style style:name="C" style:family="text"><style:text-properties fo:font-size="6pt"/></style:style>"#,
            body: r#"<text:list text:style-name="L"><text:list-item><text:p>a</text:p></text:list-item></text:list>"#,
            settings: LIST_SETTINGS,
            ..Fixture::default()
        })
        .unwrap();
        let DocRun::Text(label) = &paragraphs(&document)[0].runs[0] else {
            panic!("label run")
        };
        assert_eq!(label.font_size, 20.0);
    }

    #[test]
    fn places_list_paragraphs_as_libreoffice() {
        let styles = list_styles(&list_level(1, "1", "", 2.0, -1.0, 3.0));
        let body = r#"<text:list text:style-name="L"><text:list-item><text:p>a</text:p><text:p>b</text:p></text:list-item></text:list>"#;
        let cm = 72.0 / 2.54;
        let tabs = |relative: bool| {
            let document = read_fixture(Fixture {
                styles: &styles,
                body,
                settings: &format!(r#"<config:config-item config:name="TabsRelativeToIndent" config:type="boolean">{relative}</config:config-item>"#),
                ..Fixture::default()
            })
            .unwrap();
            let paragraphs = paragraphs(&document);
            let labelled = paragraphs[0];
            assert!((labelled.indent_left - 2.0 * cm).abs() < 1e-9);
            assert!((labelled.indent_first + cm).abs() < 1e-9);
            assert!(labelled.no_hanging_indent_tab);
            assert_eq!(texts(labelled)[..2], ["1.".to_string(), "\t".to_string()]);
            // A later paragraph of the item sits at M.
            assert!((paragraphs[1].indent_left - 2.0 * cm).abs() < 1e-9);
            assert_eq!(paragraphs[1].indent_first, 0.0);
            labelled
                .tab_stops
                .iter()
                .take(3)
                .map(|stop| (stop.pos / cm * 100.0).round() / 100.0)
                .collect::<Vec<_>>()
        };
        // T, then the 1.25 cm grid beyond it: from M when tabs are relative to
        // the indent, from the text area otherwise (measured).
        assert_eq!(tabs(true), vec![3.0, 3.25, 4.5]);
        assert_eq!(tabs(false), vec![3.0, 3.75, 5.0]);
    }

    #[test]
    fn rejects_unmodelled_lists() {
        let level = list_level(1, "1", "", 2.0, -1.0, 2.0);
        let styles = list_styles(&(level.clone() + &list_level(2, "1", "", 3.0, -1.0, 3.0)));
        let list = |items: &str| format!(r#"<text:list text:style-name="L">{items}</text:list>"#);
        let cases: Vec<(String, String, &str)> = vec![
            // A second sublist in one item restarts its level in LibreOffice.
            (styles.clone(), list("<text:list-item><text:p>a</text:p><text:list><text:list-item><text:p>b</text:p></text:list-item></text:list><text:list><text:list-item><text:p>c</text:p></text:list-item></text:list></text:list-item>"), "list-sublists"),
            (styles.clone(), list("<text:list-item><text:h>a</text:h></text:list-item>"), "list-heading"),
            (styles.clone(), list("<text:list-item><text:p/></text:list-item>"), "list-empty-label"),
            (styles.clone(), r#"<text:list><text:list-item><text:p>a</text:p></text:list-item></text:list>"#.to_string(), "list-style-missing"),
            (list_styles(&level.replace(
                    r#"space-mode="label-alignment""#,
                    r#"space-mode="label-width-and-position""#,
                )), list("<text:list-item><text:p>a</text:p></text:list-item>"), "list-legacy-mode"),
            (list_styles(&level.replace(r#"style:num-format="1""#, r#"style:num-format="01, 02, 03, ...""#)), list("<text:list-item><text:p>a</text:p></text:list-item>"), "list-format"),
            (list_styles(&level.replace(r#"style:num-suffix=".""#, "style:num-suffix=\"\u{a0}\"")), list("<text:list-item><text:p>a</text:p></text:list-item>"), "list-affix"),
            (styles.clone(), list(r#"<text:list-item text:start-value="32768"><text:p>a</text:p></text:list-item>"#), "list-start-value"),
            (list_styles(&level.replace("fo:margin-left=\"2cm\"", "fo:margin-left=\"30cm\"")), list("<text:list-item><text:p>a</text:p></text:list-item>"), "list-geometry"),
            (list_styles(&level.replace("<text:list-level-style-number", r#"<text:list-level-style-image text:level="1"/><text:list-level-style-number"#).replace(r#"text:level="1" style:num"#, r#"text:level="2" style:num"#)), list("<text:list-item><text:p>a</text:p></text:list-item>"), "list-image"),
        ];
        for (styles, body, code) in cases {
            assert_eq!(
                read_list(&styles, &body).unwrap_err(),
                unsupported(code),
                "{body}"
            );
        }
        let one = |styles: &str| {
            read_list(
                styles,
                &list("<text:list-item><text:p>a</text:p></text:list-item>"),
            )
        };
        // A circled-number label is Asian script, which the reader rejects.
        assert_eq!(
            one(&list_styles(&level.replace(
                r#"style:num-format="1""#,
                r#"style:num-format="①, ②, ③, ...""#
            )))
            .unwrap_err(),
            unsupported("list-label-script")
        );
        // A U+2011 bullet is drawn as a gap by LibreOffice.
        assert_eq!(
            one(r#"<text:list-style style:name="L"><text:list-level-style-bullet text:level="1" text:bullet-char="&#x2011;"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2cm" fo:text-indent="-1cm" fo:margin-left="2cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>"#).unwrap_err(),
            unsupported("list-bullet")
        );
        // A header paragraph beyond the text area.
        assert_eq!(
            read_list(
                &list_styles(&level.replace(r#"fo:margin-left="2cm""#, r#"fo:margin-left="30cm""#)),
                &list("<text:list-header><text:p>h</text:p></text:list-header>"),
            )
            .unwrap_err(),
            unsupported("list-geometry")
        );
        // Line separators and format characters in an affix.
        for c in [
            "&#x2028;", "&#x2029;", "&#x200B;", "&#x2060;", "%1%", "&#x301;", "&#xFE0F;",
        ] {
            assert_eq!(
                one(&list_styles(&level.replace(
                    r#"style:num-suffix=".""#,
                    &format!(r#"style:num-suffix="{c}.""#)
                )))
                .unwrap_err(),
                unsupported("list-affix"),
                "{c}"
            );
        }
        // A label that could wrap.
        assert_eq!(
            one(&list_styles(&level.replace(
                r#"style:num-suffix=".""#,
                &format!(
                    r#"style:num-prefix="{}" style:num-suffix=".""#,
                    "X".repeat(30)
                )
            )))
            .unwrap_err(),
            unsupported("list-label-extent")
        );
        // A list tab stop before the text area.
        assert_eq!(
            one(&list_styles(&level.replace(
                r#"list-tab-stop-position="2cm""#,
                r#"list-tab-stop-position="-1cm""#
            )))
            .unwrap_err(),
            unsupported("list-tab-position")
        );
        // Repeated property groups in a level.
        assert!(one(&list_styles(&level.replace(
            "</text:list-level-style-number>",
            r#"<style:text-properties/><style:text-properties/></text:list-level-style-number>"#,
        )))
        .is_err());
        // Repeated level definitions, also spelled differently.
        assert_eq!(
            one(&list_styles(
                &(level.clone() + &level.replace(r#"text:level="1""#, r#"text:level="01""#))
            ))
            .unwrap_err(),
            unsupported("list-level-duplicate")
        );
        assert_eq!(
            one(&list_styles(&(level.clone() + &level))).unwrap_err(),
            unsupported("list-level-duplicate")
        );
        // An end-aligned label that could start before the page edge; a short
        // one is admitted.
        let end = |prefix: &str| {
            list_styles(
                &level
                    .replace(
                        r#"space-mode="label-alignment""#,
                        r#"space-mode="label-alignment" fo:text-align="end""#,
                    )
                    .replace(
                        r#"style:num-suffix=".""#,
                        &format!(r#"style:num-prefix="{prefix}" style:num-suffix=".""#),
                    ),
            )
        };
        assert!(one(&end("")).is_ok());
        assert_eq!(
            one(&end(&"X".repeat(12))).unwrap_err(),
            unsupported("list-label-extent")
        );
        // ... or that could wrap from M + I to the line end.
        let far = end("")
            .replace(r#"fo:margin-left="2cm""#, r#"fo:margin-left="15.5cm""#)
            .replace(r#"fo:text-indent="-1cm""#, r#"fo:text-indent="0cm""#)
            .replace(
                r#"list-tab-stop-position="2cm""#,
                r#"list-tab-stop-position="15.7cm""#,
            );
        assert_eq!(one(&far).unwrap_err(), unsupported("list-label-extent"));
        // `continue-numbering` with nothing to continue restarts.
        assert!(read_list(&styles, r#"<text:list text:style-name="L" text:continue-numbering="true"><text:list-item><text:p>a</text:p></text:list-item></text:list>"#).is_ok());
        // Continuing a list that counted nothing, or a nested list.
        assert_eq!(
            read_list(&styles, &(list(r#"<text:list-header><text:p>h</text:p></text:list-header>"#) + r#"<text:list text:style-name="L" text:continue-numbering="true"><text:list-item><text:p>a</text:p></text:list-item></text:list>"#)).unwrap_err(),
            unsupported("list-continuation")
        );
        // Centre-aligned labelled paragraphs.
        let centred = format!(
            r#"{styles}<style:style style:name="C" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>"#
        );
        assert_eq!(
            read_list(
                &centred,
                &list(r#"<text:list-item><text:p text:style-name="C">a</text:p></text:list-item>"#)
            )
            .unwrap_err(),
            unsupported("list-paragraph-alignment")
        );
        // Grid stops need TabsRelativeToIndent.
        assert_eq!(
            read_fixture(Fixture {
                styles: &styles,
                body: &list("<text:list-item><text:p>a</text:p></text:list-item>"),
                ..Fixture::default()
            })
            .unwrap_err(),
            unsupported("list-tabs-relative-unset")
        );
    }
}
