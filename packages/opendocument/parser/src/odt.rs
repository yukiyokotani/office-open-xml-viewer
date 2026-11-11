//! OpenDocument text (ODT) → DOCX `Document`.
//!
//! The reader renders every ODF 1.3 text feature it meets by mapping it into
//! the DOCX document model, following the specification (Part 3) and, where
//! the specification leaves a choice or LibreOffice reads it differently,
//! LibreOffice as the reference consumer. It rejects only input that is
//! corrupt, encrypted, unsafe or beyond the resource limits below; a feature
//! the DOCX model cannot express exactly is mapped to its nearest DOCX
//! equivalent rather than refused (each such approximation is noted beside
//! its mapping).
//!
//! Module map: `odt_text` (inline content, white space and character
//! properties), `odt_page` (page layouts, master pages, headers and footers,
//! sections), `odt_list` (lists and heading numbering), `odt_table`
//! (tables), `odt_draw` (frames, pictures, text boxes and shapes),
//! `odt_geometry` (custom-shape geometry) and `odt_image` (picture sniffing).

use crate::error::invalid;
use crate::iri;
use crate::package::OdfPackage;
use crate::script::ScriptClass;
use crate::style::{
    Cascade, Family, Group, Scope, Style, StyleRef, StyleSheet, LOEXT_GRAPHIC, PARAGRAPH, TEXT,
};
use crate::xml::{
    self, NS_DRAW, NS_FO, NS_LOEXT, NS_OFFICE, NS_STYLE, NS_SVG, NS_TABLE, NS_TEXT, NS_XLINK,
};
use docx_model::owned_cursor::OwnedDocumentResource;
use docx_model::{
    BodyElement, BreakType, CellElement, DocParagraph, DocRun, DocTable, Document,
    DocumentSettings, DocxNote, HeadersFooters, LineSpacing, ParaBorderEdge, ParagraphBorders,
    RunFontFacts, TabStop, TextRun,
};
use std::cell::{Cell, RefCell};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::rc::Rc;

#[path = "odt_draw.rs"]
mod draw;
#[path = "odt_geometry.rs"]
mod geometry;
#[path = "odt_image.rs"]
mod image;
#[path = "odt_list.rs"]
mod list;
#[path = "odt_math.rs"]
mod math;
#[path = "odt_page.rs"]
mod page;
#[path = "odt_table.rs"]
mod table;
#[path = "odt_text.rs"]
mod text;

use text::{FontFaces, Inline, TextContext};

/// Media types an ODT source claims (Part 2 §3.3; templates share the model).
pub const MEDIA_TYPES: [&str; 2] = [
    "application/vnd.oasis.opendocument.text",
    "application/vnd.oasis.opendocument.text-template",
];

/// Implementation bound on nested inline elements walked recursively.
const MAX_INLINE_DEPTH: usize = 32;
/// Implementation bound on nested block containers (sections, lists, tables,
/// frames, notes).
const MAX_BLOCK_DEPTH: usize = 32;
/// Implementation bound on retained text, the dominant construction cost of
/// this reader; the DOCX cursor additionally enforces its JSON ceilings.
const MAX_RETAINED_TEXT_BYTES: usize = 64 * 1024 * 1024;
/// Implementation bound on paragraphs.
const MAX_PARAGRAPHS: usize = 1_000_000;
/// Implementation bound on the characters of one paragraph (after `text:s`
/// expansion). It bounds the transient token buffer of §6.1.2 normalization
/// independently of the document-wide text budget.
const MAX_PARAGRAPH_CHARS: usize = 1_000_000;
/// Implementation bound on distinct span contexts in one paragraph.
const MAX_PARAGRAPH_CONTEXTS: usize = 65_536;
/// ODT-owned construction policy (a library policy, not a format limit):
/// the estimated native size of the retained model (text, runs, paragraphs,
/// pictures), charged before each allocation. The DOCX cursor additionally
/// enforces its retained JSON ceiling while serializing.
const MAX_CONSTRUCTION_BYTES: usize = 384 * 1024 * 1024;
/// Implementation bound on the §16.2 search order of one span: its own style
/// chain and those of the enclosing spans and paragraph.
const MAX_SEARCH_ORDER: usize = 256;
/// Implementation bound on cascade work, document-wide: every paragraph and
/// run resolution visits each style of its search order and each attribute
/// of the resolved property group (one unit each plus one per byte).
const MAX_CASCADE_WORK: usize = 256 * 1024 * 1024;
/// Estimated native size of one table cell and of one empty paragraph in it.
const CELL_BYTES: usize = std::mem::size_of::<docx_model::DocTableCell>() + 256;
const EMPTY_PARAGRAPH_BYTES: usize =
    std::mem::size_of::<DocParagraph>() + std::mem::size_of::<CellElement>();

/// The admitted document and the package files its runs reference by key.
pub fn read(package: &OdfPackage) -> Result<(Document, Vec<OwnedDocumentResource>), String> {
    let content_text = package.read_text("content.xml")?;
    // ODF requires styles.xml in a package (Part 3 §3.1.3); a package
    // without it is read with no common styles.
    let styles_text = if package.contains("styles.xml") {
        package.read_text("styles.xml")?
    } else {
        String::from(
            r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
        )
    };
    let styles_doc = xml::parse(&styles_text)?;
    let content_doc = xml::parse(&content_text)?;
    let styles_root = styles_doc.root_element();
    let content_root = content_doc.root_element();
    if !xml::is(styles_root, NS_OFFICE, "document-styles")
        || !xml::is(content_root, NS_OFFICE, "document-content")
    {
        return Err(invalid("styles.xml or content.xml has an unexpected root"));
    }
    let version = content_root.attribute((NS_OFFICE, "version"));

    let mut sheet = StyleSheet::default();
    let mut fonts = FontFaces::default();
    for (root, scope) in [(styles_root, Scope::Styles), (content_root, Scope::Content)] {
        for child in root.children().filter(|n| n.is_element()) {
            match (child.tag_name().namespace(), child.tag_name().name()) {
                (Some(NS_OFFICE), "font-face-decls") => fonts.add(child, package)?,
                (Some(NS_OFFICE), "styles") => {
                    sheet.add_common(child)?;
                    sheet.add_outline(child)?;
                }
                (Some(NS_OFFICE), "automatic-styles") => sheet.add_automatic(scope, child)?,
                _ => {}
            }
        }
    }
    let settings = if package.contains("settings.xml") {
        text::read_settings(&package.read_text("settings.xml")?)?
    } else {
        text::LayoutFlags::default()
    };
    let common_styles = styles_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "styles"));
    let pages = page::Pages::read(styles_root)?;
    let body = content_root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|b| b.children().find(|n| xml::is(*n, NS_OFFICE, "text")))
        .ok_or_else(|| invalid("content.xml has no office:text body"))?;

    let default_tab_stop = sheet
        .default_style(Family::Paragraph)
        .and_then(|style| Cascade::new(vec![style], PARAGRAPH).get(NS_STYLE, "tab-stop-distance"))
        .and_then(|value| length_pt(value).ok())
        .filter(|value| *value > 0.0);
    let notes = page::NotesConfiguration::read(common_styles);
    let line_numbering = page::line_numbering(common_styles);
    let changes = text::TrackedChanges::read(body);
    let defs = draw::Definitions::read(common_styles);

    let initial = pages.initial_master(body);
    let geometry = pages.geometry(initial);
    let mut reader = Reader {
        package,
        sheet: &sheet,
        fonts: &fonts,
        pages: &pages,
        notes_config: &notes,
        changes: &changes,
        defs: &defs,
        scope: Scope::Content,
        settings,
        images: Vec::new(),
        image_types: HashMap::new(),
        lists: list::Lists::default(),
        list_paragraph: None,
        text_width: geometry.text_width,
        text_height: geometry.text_height,
        page_left_margin: geometry.left,
        page_right_margin: geometry.right,
        story: Story::Body,
        blocks: Vec::new(),
        pending_anchors: Vec::new(),
        section_indent: (0.0, 0.0),
        current_master: initial.to_string(),
        sections: Vec::new(),
        area: 0,
        next_area: 0,
        depth: 0,
        footnotes: Vec::new(),
        endnotes: Vec::new(),
        next_note: 0,
        next_occurrence: 0,
        next_table: 0,
        headers: HashMap::new(),
        used_faces: RefCell::new(HashSet::new()),
        chains: HashMap::new(),
        cascade_work: Cell::new(0),
        paragraph_count: 0,
        retained_text_bytes: 0,
        construction_bytes: 0,
        resource_bytes: 0,
        contextual_spacing_defined: version != Some("1.2"),
    };
    // The first section: the initial master page, with its page layout's
    // columns (§17.12) like every later page-style section.
    let first = page::SectionStart {
        master: initial.to_string(),
        kind: "nextPage",
        columns: pages.columns(initial),
        page_start: None,
        block: 0,
    };
    reader.charge_construction(page::section_bytes(&first))?;
    reader.sections.push(first);
    reader.walk_blocks(body)?;
    let blocks = std::mem::take(&mut reader.blocks);
    let (body, section_marks) = reader.finish_body(blocks)?;
    let mut document = reader.assemble(body, section_marks, line_numbering)?;

    // The font table and its metadata maps copy each used family name.
    let font_table_bytes = reader
        .used_faces
        .borrow()
        .iter()
        .map(|face| 4 * (face.family.len() + 64))
        .sum::<usize>();
    reader.charge_construction(font_table_bytes)?;
    let used = text::used_font_table(&reader.used_faces.borrow());
    document.font_family_classes = text::metadata(&used, |face| face.generic);
    document.font_family_pitches = text::metadata(&used, |face| face.pitch);
    document.font_family_charsets = text::metadata(&used, |face| face.charset);
    document.settings = Some(DocumentSettings {
        default_tab_stop,
        // DoNotJustifyLinesWithManualBreak leaves a justified line ended by a
        // manual break unjustified, as DOCX doNotExpandShiftReturn does
        // (§17.15.3.5).
        do_not_expand_shift_return: reader.settings.no_justify_breaks.then_some(true),
        // LibreOffice gives each column its own footnote area: a note sits
        // under the column holding its reference and continues at the top of
        // the next column's area.
        footnote_column_bands: true,
        // LibreOffice counts grid cells for a snapped line from the face's
        // own box, without Word's Far East 1.3 x usWin rule (measured:
        // MS Mincho and Meiryo grid controls; unsnapped lines keep 1.3),
        // and admits a snapped line at the page bottom only when its whole
        // cells fit (rr3 family D: 文 stays at 16.4pt left, moves at 16.35pt).
        plain_font_line_boxes: true,
        ..DocumentSettings::default()
    });
    document.footnotes = std::mem::take(&mut reader.footnotes);
    document.endnotes = std::mem::take(&mut reader.endnotes);
    document.note_layout_settings = notes.wire(&pages, &reader.current_master);
    reader.check_budgets()?;
    let images = std::mem::take(&mut reader.images);
    Ok((document, images))
}

/// Which story the reader is filling: the body owns sections and page
/// breaks; other stories (headers, notes, cells, text boxes) flow inside a
/// container.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Story {
    Body,
    Contained,
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
    /// The content area (body, `text:section`, cell, frame) holding the
    /// paragraph: §20.255 suppresses spacing only within one content area.
    area: u32,
    page_break_after: bool,
    column_break_after: bool,
}

/// A table with the vertical margins it keeps outside its rows (§20.204 et
/// al. on `style:table-properties`).
struct TableBlock {
    table: DocTable,
    top: f64,
    bottom: f64,
    page_break_before: bool,
    page_break_after: bool,
}

enum Block {
    Paragraph(Box<Collected>),
    Table(Box<TableBlock>),
    ColumnBreak,
    /// The start of a new DOCX section (index into `Reader::sections`).
    Section(usize),
}

struct Reader<'a> {
    package: &'a OdfPackage,
    sheet: &'a StyleSheet,
    fonts: &'a FontFaces,
    pages: &'a page::Pages<'a, 'a>,
    notes_config: &'a page::NotesConfiguration,
    changes: &'a text::TrackedChanges<'a, 'a>,
    defs: &'a draw::Definitions<'a>,
    /// Automatic styles of content.xml for body content, of styles.xml for
    /// master-page content (§3.15).
    scope: Scope,
    settings: text::LayoutFlags,
    /// Package files referenced by runs, keyed by their package path, and
    /// each one's media type.
    images: Vec<OwnedDocumentResource>,
    image_types: HashMap<String, &'static str>,
    lists: list::Lists,
    list_paragraph: Option<list::ListParagraph<'a>>,
    /// The current page's text area (pt) and side margins.
    text_width: f64,
    text_height: f64,
    page_left_margin: f64,
    page_right_margin: f64,
    story: Story,
    blocks: Vec<Block>,
    /// Page-anchored frames waiting for the next paragraph to carry them.
    pending_anchors: Vec<DocRun>,
    /// Indents of the enclosing `text:section`s (§20.204 on section
    /// properties), added to their paragraphs.
    section_indent: (f64, f64),
    current_master: String,
    sections: Vec<page::SectionStart>,
    area: u32,
    next_area: u32,
    depth: usize,
    footnotes: Vec<DocxNote>,
    endnotes: Vec<DocxNote>,
    next_note: u32,
    next_occurrence: u32,
    next_table: u32,
    /// Built header and footer stories per master page.
    headers: HashMap<String, Rc<page::MasterStories>>,
    used_faces: RefCell<HashSet<text::Face>>,
    /// Each applied style's parent chain, by style identity.
    chains: HashMap<*const Style, Rc<[StyleRef<'a>]>>,
    cascade_work: Cell<usize>,
    paragraph_count: usize,
    retained_text_bytes: usize,
    construction_bytes: usize,
    /// The part of `construction_bytes` charged for package resources
    /// (pictures), which runs reference by key and so never copy.
    resource_bytes: usize,
    contextual_spacing_defined: bool,
}

impl<'a> Reader<'a> {
    fn enter(&mut self) -> Result<(), String> {
        self.depth += 1;
        if self.depth > MAX_BLOCK_DEPTH {
            return Err("OpenDocument content nests beyond the reader's limit".into());
        }
        Ok(())
    }

    fn leave(&mut self) {
        self.depth -= 1;
    }

    /// Block content (§5): paragraphs, headings, lists, tables, sections,
    /// indexes and frames anchored to the page.
    fn walk_blocks(&mut self, parent: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        self.enter()?;
        for node in parent.children().filter(|n| n.is_element()) {
            self.block(node)?;
        }
        self.leave();
        Ok(())
    }

    fn block(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        match (node.tag_name().namespace(), node.tag_name().name()) {
            (Some(NS_TEXT), "p") => self.paragraph(node, None)?,
            (Some(NS_TEXT), "h") => {
                let level = node
                    .attribute((NS_TEXT, "outline-level"))
                    .and_then(|v| v.parse::<u32>().ok())
                    .filter(|v| (1..=10).contains(v))
                    .unwrap_or(1);
                self.heading(node, level)?
            }
            (Some(NS_TEXT), "list") => self.list(node, None)?,
            (Some(NS_TEXT), "numbered-paragraph") => self.numbered_paragraph(node)?,
            (Some(NS_TABLE), "table") => self.table(node)?,
            (Some(NS_TEXT), "section") => self.section(node)?,
            // Indexes (§8): the index body holds the generated entries as
            // ordinary paragraphs (and its title as a section-like block).
            (
                Some(NS_TEXT),
                "table-of-content" | "illustration-index" | "table-index" | "object-index"
                | "user-index" | "alphabetical-index" | "bibliography",
            ) => {
                if let Some(index_body) =
                    node.children().find(|n| xml::is(*n, NS_TEXT, "index-body"))
                {
                    self.index_body(index_body)?;
                }
            }
            (Some(NS_TEXT), "index-title") => self.index_body(node)?,
            // Frames anchored to the page or paragraph at body level
            // (§9.3, `text:anchor-type="page"`): carried by the next
            // paragraph.
            (Some(NS_DRAW), _) => {
                if let Some(run) = self.drawing(node, None)? {
                    self.pending_anchors.push(run);
                }
            }
            // A producer hint (§5.6), declarations without rendering, forms
            // (controls are drawn by their draw:control shapes) and the
            // tracked-change log (read up front).
            _ => {}
        }
        Ok(())
    }

    fn index_body(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        self.enter()?;
        for child in node.children().filter(|n| n.is_element()) {
            self.block(child)?;
        }
        self.leave();
        Ok(())
    }

    /// `text:numbered-paragraph` (§5.3.6): one numbered paragraph outside a
    /// list element, numbered as a list item of its list at its level.
    fn numbered_paragraph(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        let Some(child) = node
            .children()
            .find(|n| xml::is(*n, NS_TEXT, "p") || xml::is(*n, NS_TEXT, "h"))
        else {
            return Ok(());
        };
        self.numbered(node, child)
    }

    /// A heading (§5.1.2), numbered by the outline style (§16.36) unless it
    /// is a list header or sits in a list.
    fn heading(&mut self, node: roxmltree::Node<'a, 'a>, level: u32) -> Result<(), String> {
        if self.list_paragraph.is_none()
            && node.attribute((NS_TEXT, "is-list-header")) != Some("true")
        {
            // A paragraph style naming a list style (§19.509
            // style:list-style-name) overrides the outline numbering: an
            // empty name removes it (LibreOffice).
            let chain = self.paragraph_chain(node.attribute((NS_TEXT, "style-name")));
            let link = chain.iter().find_map(|style| {
                style
                    .style()
                    .attributes
                    .get(&(NS_STYLE.to_string(), "list-style-name".to_string()))
            });
            match link.map(String::as_str) {
                Some("") => {}
                Some(name) if self.sheet.find_list(self.scope, name).is_some() => {
                    self.list_paragraph = self.styled_list_paragraph(node, name, level)?;
                }
                _ => self.list_paragraph = self.outline_paragraph(node, level)?,
            }
        }
        self.paragraph(node, Some(level))
    }

    /// `text:section` (§5.4): columns and indents for its content, or no
    /// content when it is hidden.
    fn section(&mut self, node: roxmltree::Node<'a, 'a>) -> Result<(), String> {
        if node.attribute((NS_TEXT, "display")) == Some("none")
            || node.attribute((NS_TEXT, "display")) == Some("condition")
                && node.attribute((NS_TEXT, "condition")).is_some()
        {
            return Ok(());
        }
        let style = node
            .attribute((NS_TEXT, "style-name"))
            .and_then(|name| self.sheet.find(self.scope, Family::Section, name));
        let properties = style.and_then(|s| s.group("section-properties"));
        let indent = properties.map_or((0.0, 0.0), |p| {
            let get = |name: &str| {
                p.attributes
                    .get(&(NS_FO.to_string(), name.to_string()))
                    .and_then(|v| length_pt(v).ok())
                    .unwrap_or(0.0)
            };
            (get("margin-left"), get("margin-right"))
        });
        let columns = match properties
            .and_then(|p| p.elements.get("columns"))
            .and_then(|elements| elements.first())
        {
            Some(columns) => page::columns(columns, self.text_width - indent.0 - indent.1)?,
            None => None,
        };
        let enclosing_indent = self.section_indent;
        self.section_indent = (enclosing_indent.0 + indent.0, enclosing_indent.1 + indent.1);
        // A section with its own columns is a continuous DOCX section; the
        // content after it continues in the enclosing columns.
        let sectioned = self.story == Story::Body && columns.is_some();
        let enclosing_columns = self.sections.last().and_then(|s| s.columns.clone());
        if sectioned {
            self.start_section(page::SectionStart {
                master: self.current_master.clone(),
                kind: "continuous",
                columns: columns.clone(),
                page_start: None,
                block: 0,
            })?;
        }
        let enclosing_area = self.area;
        self.next_area += 1;
        self.area = self.next_area;
        self.walk_blocks(node)?;
        self.area = enclosing_area;
        self.section_indent = enclosing_indent;
        // §20.336 text:dont-balance-text-columns: an unbalanced section that
        // ends the document stays the last DOCX section, which the layout
        // does not balance (a continuous section after it would).
        let unbalanced_end = properties
            .and_then(|p| {
                p.attributes
                    .get(&(NS_TEXT.to_string(), "dont-balance-text-columns".to_string()))
            })
            .is_some_and(|v| v == "true")
            && node.next_siblings().skip(1).all(|n| !n.is_element())
            && node.parent().is_some_and(|p| xml::is(p, NS_OFFICE, "text"));
        if sectioned && !unbalanced_end {
            self.start_section(page::SectionStart {
                master: self.current_master.clone(),
                kind: "continuous",
                columns: enclosing_columns,
                page_start: None,
                block: 0,
            })?;
        }
        Ok(())
    }

    /// Records a section start. Every section later becomes a DOCX section
    /// break whose properties, columns and header/footer stories are copies;
    /// the start and its fixed-size copies are charged here, the stories
    /// when they are copied (`section_parts`).
    fn start_section(&mut self, mut start: page::SectionStart) -> Result<(), String> {
        self.charge_construction(page::section_bytes(&start))?;
        start.block = self.blocks.len();
        self.sections.push(start);
        self.blocks.push(Block::Section(self.sections.len() - 1));
        Ok(())
    }

    /// Read `node`'s block content into a separate story (cell, frame, note,
    /// header), returning its body elements.
    fn contained_story(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        width: f64,
    ) -> Result<Vec<BodyElement>, String> {
        let saved_blocks = std::mem::take(&mut self.blocks);
        let saved_story = std::mem::replace(&mut self.story, Story::Contained);
        let saved_width = std::mem::replace(&mut self.text_width, width.max(1.0));
        let saved_indent = std::mem::replace(&mut self.section_indent, (0.0, 0.0));
        let saved_anchors = std::mem::take(&mut self.pending_anchors);
        let saved_list = self.list_paragraph.take();
        let enclosing_area = self.area;
        self.next_area += 1;
        self.area = self.next_area;
        let walked = self.walk_blocks(node);
        self.area = enclosing_area;
        let blocks = std::mem::replace(&mut self.blocks, saved_blocks);
        self.story = saved_story;
        self.text_width = saved_width;
        self.section_indent = saved_indent;
        // Frames anchored inside the story without a following paragraph.
        let orphans = std::mem::replace(&mut self.pending_anchors, saved_anchors);
        self.list_paragraph = saved_list;
        walked?;
        let mut elements = self.finish_story(blocks)?;
        if !orphans.is_empty() {
            elements.push(BodyElement::Paragraph(Box::new(anchor_paragraph(orphans))));
        }
        Ok(elements)
    }

    /// A paragraph or heading (§5.1).
    fn paragraph(
        &mut self,
        node: roxmltree::Node<'a, 'a>,
        outline_level: Option<u32>,
    ) -> Result<(), String> {
        self.count_paragraph()?;
        let list = self.list_paragraph.take();
        // §7.7.11, §19.821: a hidden-paragraph field whose condition last
        // evaluated true hides its whole paragraph.
        if text::hidden_by_field(node) {
            return Ok(());
        }
        let chain = self.paragraph_chain(node.attribute((NS_TEXT, "style-name")));
        let named_style = chain
            .iter()
            .find(|s| !s.style().automatic)
            .map(|s| s.style() as *const Style);
        self.charge_cascade(&chain, PARAGRAPH)?;
        let properties = Cascade::new(chain.clone(), PARAGRAPH);

        // §19.501: a paragraph whose style names a master page starts a page
        // with that page style (a page number restarts the count there).
        let master = chain.iter().find_map(|s| {
            s.style()
                .attributes
                .get(&(NS_STYLE.to_string(), "master-page-name".to_string()))
                .map(String::as_str)
        });
        let mut page_break_before = false;
        if let Some(master) = master.filter(|name| !name.is_empty()) {
            if self.story == Story::Body && self.pages.has_master(master) {
                let page_start = properties
                    .get(NS_STYLE, "page-number")
                    .and_then(|v| v.parse::<i64>().ok());
                if self.has_body_content() {
                    let kind = self.pages.start_kind(master);
                    self.current_master = master.to_string();
                    self.start_section(page::SectionStart {
                        master: master.to_string(),
                        kind,
                        columns: self.pages.columns(master),
                        page_start,
                        block: 0,
                    })?;
                } else {
                    // The first paragraph names the initial page style.
                    self.current_master = master.to_string();
                    if let Some(first) = self.sections.last_mut() {
                        first.master = master.to_string();
                        first.columns = self.pages.columns(master);
                        first.page_start = page_start;
                    }
                }
                self.apply_page_geometry();
            } else if self.story == Story::Body {
                page_break_before = self.has_body_content();
            }
        }

        let mut collected = self.paragraph_properties(&chain, &properties)?;
        if collected.contextual && !self.contextual_spacing_defined {
            collected.contextual = false;
        }
        collected.named_style = named_style;
        collected.area = self.area;
        collected.paragraph.page_break_before |= page_break_before;
        collected.paragraph.outline_level = outline_level.map(|level| level.saturating_sub(1));
        collected.paragraph.indent_left += self.section_indent.0;
        collected.paragraph.indent_right += self.section_indent.1;
        if let Some(list) = &list {
            self.list_geometry(list, &chain, &properties, &mut collected.paragraph)?;
        }
        if collected.paragraph.tab_stops.is_empty() {
            collected.paragraph.tab_stops =
                self.tab_stops(&properties, collected.paragraph.indent_left)?;
        }

        // The paragraph's own text properties describe its mark (empty
        // paragraph height) and the text outside spans.
        let mut inline = Inline::new(chain.into());
        // The paragraph mark's formatting: LibreOffice's paragraph-marker
        // character style (`loext:marker-style-name`) over the paragraph's.
        let mark_context = match node
            .attribute((NS_LOEXT, "marker-style-name"))
            .and_then(|name| self.sheet.find(self.scope, Family::Text, name))
        {
            Some(style) => {
                let styles = self.style_chain(style);
                inline.contexts.push(TextContext {
                    styles,
                    parent: Some(0),
                    link: None,
                });
                (inline.contexts.len() - 1) as u32
            }
            None => 0,
        };
        let mark = self.run_properties(&inline.contexts, mark_context, ScriptClass::Latin)?;
        let mark_asian = self.run_properties(&inline.contexts, mark_context, ScriptClass::Asian)?;
        self.charge_construction(
            std::mem::size_of::<Collected>()
                + std::mem::size_of::<BodyElement>()
                + 5 * mark.family.len(),
        )?;
        let mut mark_facts = mark.font_facts();
        mark_facts.font_family_east_asia = Some(mark_asian.family.to_string());
        collected.paragraph.paragraph_mark_font_facts = Some(mark_facts);
        collected.paragraph.default_font_size = Some(mark.size);
        collected.paragraph.default_font_family = Some(mark.family.to_string());
        collected.paragraph.default_font_family_east_asia = Some(mark_asian.family.to_string());
        if mark.hidden {
            collected.paragraph.mark_vanish = true;
        }

        self.inline(node, 0, &mut inline, 0)?;
        self.emit_runs(&mut inline, &mut collected.paragraph)?;
        let drop_cap = self.drop_cap(&properties, &mut collected.paragraph)?;
        if let Some(label) = list.and_then(|list| list.label) {
            self.list_label(label, &mut inline.contexts, &mut collected.paragraph)?;
        }
        // Anchored frames whose anchor is this paragraph's page or an
        // earlier block.
        if !self.pending_anchors.is_empty() {
            let anchors = std::mem::take(&mut self.pending_anchors);
            // Each anchor gains a host run copying the mark's faces.
            self.charge_construction(anchors.len().saturating_mul(mark.retained_estimate()))?;
            let host = mark.text_run();
            let anchors: Vec<DocRun> = anchors
                .into_iter()
                .flat_map(|run| text::with_anchor_host(run, &host))
                .collect();
            collected.paragraph.runs.splice(0..0, anchors);
        }
        if let Some(shading) = collected.paragraph.shading.clone() {
            // §20.395 automatic colour (also an unset colour) over the
            // paragraph background, resolved as LibreOffice paints it.
            let resolved = text::libreoffice_automatic_color(&shading);
            for run in &mut collected.paragraph.runs {
                if let DocRun::Text(text) = run {
                    if text.color.is_none() && text.background.is_none() {
                        text.color = Some(resolved.to_string());
                        text.color_auto = false;
                    }
                }
            }
        }
        if let Some(cap) = drop_cap {
            // The cap's framed paragraph precedes its paragraph without
            // spacing of its own (DOCX drop caps, §17.3.1.11).
            let cap = Collected {
                paragraph: cap,
                top: collected.top,
                bottom: 0.0,
                contextual: false,
                named_style: None,
                area: collected.area,
                page_break_after: false,
                column_break_after: false,
            };
            collected.top = 0.0;
            collected.paragraph.page_break_before = false;
            self.blocks.push(Block::Paragraph(Box::new(cap)));
        }
        self.blocks.push(Block::Paragraph(Box::new(collected)));
        Ok(())
    }

    /// `style:drop-cap` (§17.9): the paragraph's first characters (or first
    /// word) set over `style:lines` lines, as a DOCX drop-cap frame
    /// (§17.3.1.11 `w:framePr w:dropCap="drop"`). The cap's size spans the
    /// lines at the paragraph's nominal line height (1.15 em; the reader has
    /// no font metrics), and `style:distance` separates it from the text.
    fn drop_cap(
        &mut self,
        properties: &Cascade<'a>,
        paragraph: &mut DocParagraph,
    ) -> Result<Option<DocParagraph>, String> {
        let Some(cap) = properties.element("drop-cap").and_then(|e| e.first()) else {
            return Ok(None);
        };
        let attribute = |name: &str| {
            cap.attributes
                .get(&(NS_STYLE.to_string(), name.to_string()))
                .map(String::as_str)
        };
        let lines = attribute("lines")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1);
        if lines < 2 {
            return Ok(None);
        }
        let first_text: String = paragraph
            .runs
            .iter()
            .take_while(|run| matches!(run, DocRun::Text(text) if text.list_label.is_none()))
            .filter_map(|run| match run {
                DocRun::Text(text) => Some(text.text.as_str()),
                _ => None,
            })
            .collect();
        let count = match attribute("length") {
            Some("word") => first_text
                .chars()
                .take_while(|c| !c.is_whitespace())
                .count(),
            Some(value) => value.parse::<usize>().unwrap_or(1),
            None => 1,
        }
        .min(first_text.chars().count());
        if count == 0 {
            return Ok(None);
        }
        // Take the first `count` characters from the leading text runs.
        let mut remaining = count;
        let mut cap_runs = Vec::new();
        let index = 0;
        while remaining > 0 && index < paragraph.runs.len() {
            let DocRun::Text(text) = &mut paragraph.runs[index] else {
                break;
            };
            let chars = text.text.chars().count();
            if chars <= remaining {
                remaining -= chars;
                cap_runs.push(paragraph.runs.remove(index));
                continue;
            }
            // The run splits in two: the text moves (no copy of it), the
            // second run's properties are a charged copy.
            self.charge_construction(text::run_copy_bytes(text))?;
            let original = std::mem::take(&mut text.text);
            let split: String = original.chars().take(remaining).collect();
            let rest: String = original.chars().skip(remaining).collect();
            drop(original);
            let mut head = text.as_ref().clone();
            head.text = split;
            text.text = rest;
            cap_runs.push(DocRun::Text(Box::new(head)));
            remaining = 0;
        }
        let size = paragraph.default_font_size.unwrap_or(12.0);
        let line = paragraph
            .line_spacing
            .as_ref()
            .map_or(size * 1.15, |spacing| match spacing.rule.as_str() {
                "auto" => size * 1.15 * spacing.value,
                _ => spacing.value,
            });
        let cap_size = line * f64::from(lines) / 1.15;
        // A character style for the cap (§19.494 style:style-name).
        let styled = attribute("style-name")
            .and_then(|name| self.sheet.find(self.scope, Family::Text, name))
            .map(|style| self.style_chain(style));
        for run in &mut cap_runs {
            if let DocRun::Text(text) = run {
                if let Some(chain) = &styled {
                    let contexts = vec![TextContext {
                        styles: chain.clone(),
                        parent: None,
                        link: None,
                    }];
                    let styled = self.run_properties(&contexts, 0, ScriptClass::Latin)?;
                    let mut replacement = styled.text_run();
                    replacement.text = std::mem::take(&mut text.text);
                    **text = replacement;
                }
                text.font_size = cap_size;
                text.font_size_cs = Some(cap_size);
            }
        }
        let distance = attribute("distance")
            .and_then(|v| length_pt(v).ok())
            .unwrap_or(0.0);
        Ok(Some(DocParagraph {
            alignment: "left".into(),
            indent_left: paragraph.indent_left,
            runs: cap_runs,
            default_font_size: Some(cap_size),
            frame_pr: Some(Box::new(docx_model::FramePr {
                drop_cap: "drop".into(),
                lines,
                wrap: "around".into(),
                h_anchor: "text".into(),
                v_anchor: "text".into(),
                h_rule: "auto".into(),
                h_space: distance,
                ..docx_model::FramePr::default()
            })),
            ..DocParagraph::default()
        }))
    }

    /// Whether the body already holds content before the current block (a
    /// page style named on the first paragraph starts no new page).
    fn has_body_content(&self) -> bool {
        self.blocks
            .iter()
            .any(|block| !matches!(block, Block::Section(_)))
    }

    /// Text area of the current master page.
    fn apply_page_geometry(&mut self) {
        let geometry = self.pages.geometry(&self.current_master);
        self.text_width = geometry.text_width;
        self.text_height = geometry.text_height;
        self.page_left_margin = geometry.left;
        self.page_right_margin = geometry.right;
    }

    /// A paragraph style's chain plus the paragraph default style.
    fn paragraph_chain(&mut self, name: Option<&str>) -> Vec<StyleRef<'a>> {
        let mut chain: Vec<StyleRef<'a>> =
            match name.and_then(|name| self.sheet.find(self.scope, Family::Paragraph, name)) {
                Some(style) => self.style_chain(style).to_vec(),
                None => Vec::new(),
            };
        if let Some(default) = self.sheet.default_style(Family::Paragraph) {
            chain.push(default);
        }
        chain
    }

    /// A style's parent chain, built once per style and charged to the
    /// cascade work budget.
    fn style_chain(&mut self, style: StyleRef<'a>) -> Rc<[StyleRef<'a>]> {
        let key: *const Style = style.style();
        if let Some(chain) = self.chains.get(&key) {
            return chain.clone();
        }
        let chain: Rc<[StyleRef<'a>]> = self
            .sheet
            .chain(style, |work| self.charge_work(work))
            .unwrap_or_else(|_| vec![style])
            .into();
        self.chains.insert(key, chain.clone());
        chain
    }

    /// Paragraph properties (§17.6) mapped onto the DOCX paragraph.
    fn paragraph_properties(
        &mut self,
        chain: &[StyleRef<'a>],
        cascade: &Cascade<'a>,
    ) -> Result<Collected, String> {
        let length = |name: &str| -> f64 {
            cascade
                .get(NS_FO, name)
                .and_then(|value| length_pt(value).ok())
                .unwrap_or(0.0)
        };
        let rtl = matches!(cascade.get(NS_STYLE, "writing-mode"), Some("rl-tb" | "rl"));
        // DOCX alignment is logical in a right-to-left paragraph (left is the
        // start edge); ODF's start/end are logical, left/right physical
        // (XSL §7.15.9).
        let alignment = match cascade.get(NS_FO, "text-align") {
            None | Some("start") => "left",
            Some("end") => "right",
            Some("left") => {
                if rtl {
                    "right"
                } else {
                    "left"
                }
            }
            Some("right") => {
                if rtl {
                    "left"
                } else {
                    "right"
                }
            }
            Some("center") => "center",
            // §20.216 fo:text-align-last: a justified last line too.
            Some("justify") if cascade.get(NS_FO, "text-align-last") == Some("justify") => {
                "distribute"
            }
            Some("justify") => "both",
            Some(_) => "left",
        };
        // Line height controls (§20.204 fo:line-height, §20.317
        // style:line-height-at-least, §20.318 style:line-spacing).
        let line_spacing = match (
            cascade.get(NS_FO, "line-height"),
            cascade.get(NS_STYLE, "line-height-at-least"),
            cascade.get(NS_STYLE, "line-spacing"),
        ) {
            (Some(value), _, _) if value.ends_with('%') => {
                // LibreOffice scales the natural line height by the
                // percentage (DOCX "auto" with the same multiple), where ODF
                // 1.4 defines it on the font size: the LibreOffice reading
                // (user-approved compatibility rule, 2026-10-31).
                percent(value)
                    .ok()
                    .filter(|v| *v > 0.0)
                    .map(|multiple| LineSpacing {
                        value: multiple,
                        rule: "auto".into(),
                        explicit: true,
                    })
            }
            (Some("normal"), _, _) | (None, None, None) => None,
            (Some(value), _, _) => {
                length_pt(value)
                    .ok()
                    .filter(|v| *v > 0.0)
                    .map(|value| LineSpacing {
                        value,
                        rule: "exact".into(),
                        explicit: true,
                    })
            }
            (None, Some(value), _) => {
                length_pt(value)
                    .ok()
                    .filter(|v| *v > 0.0)
                    .map(|value| LineSpacing {
                        value,
                        rule: "atLeast".into(),
                        explicit: true,
                    })
            }
            // Leading (extra space between lines) has no DOCX rule; it is
            // approximated as a proportional height over a 1.15 natural line
            // of the paragraph's font (not measured).
            (None, None, Some(value)) => {
                length_pt(value)
                    .ok()
                    .filter(|v| *v != 0.0)
                    .and_then(|lead| {
                        let size = text::paragraph_font_size(cascade_text(chain)).unwrap_or(12.0);
                        let natural = size * 1.15;
                        ((natural + lead) / natural > 0.0).then(|| LineSpacing {
                            value: (natural + lead) / natural,
                            rule: "auto".into(),
                            explicit: true,
                        })
                    })
            }
        };
        let count = |name: &str| -> u32 {
            cascade
                .get(NS_FO, name)
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(0)
        };
        // DOCX widow control is two lines; any ODF widow or orphan count of
        // two or more turns it on.
        let widow_control = count("widows") >= 2 || count("orphans") >= 2;
        let breaking = |name: &str| cascade.get(NS_FO, name);
        let mut indent_left = length("margin-left");
        let mut indent_right = length("margin-right");
        let mut indent_first = length("text-indent");
        // §20.256 style:auto-text-indent: a first-line indent of one em
        // (LibreOffice uses the font size).
        if cascade.get(NS_STYLE, "auto-text-indent") == Some("true") {
            indent_first = text::paragraph_font_size(cascade_text(chain)).unwrap_or(12.0);
        }
        let mut top = length("margin-top");
        let mut bottom = length("margin-bottom");
        // Relative margins (percentages of the parent style's) are resolved
        // against the parent value in the chain.
        for (name, value) in [
            ("margin-left", &mut indent_left),
            ("margin-right", &mut indent_right),
            ("text-indent", &mut indent_first),
            ("margin-top", &mut top),
            ("margin-bottom", &mut bottom),
        ] {
            if let Some(relative) = cascade.get(NS_FO, name).filter(|v| v.ends_with('%')) {
                *value = relative_length(chain, PARAGRAPH, NS_FO, name, relative);
            }
        }
        let shading = paragraph_fill(cascade, &Cascade::new(chain.to_vec(), LOEXT_GRAPHIC));
        // Borders and padding (§20.183, §20.217): a bordered side keeps its
        // padding as the border's distance from the text; LibreOffice also
        // insets an unbordered side's text by its padding, which DOCX
        // expresses as indent (left/right) or spacing (top/bottom).
        let edges = border_edges(cascade);
        let padding = padding_values(cascade);
        let mut borders = ParagraphBorders::default();
        let mut any_border = false;
        for (side, edge) in edges.into_iter().enumerate() {
            match edge {
                Some(mut edge) => {
                    any_border = true;
                    edge.space = padding[side];
                    // LibreOffice measures the paragraph indents to the
                    // border's outside; DOCX to the text.
                    match side {
                        1 => indent_left += padding[1] + edge.width,
                        3 => indent_right += padding[3] + edge.width,
                        _ => {}
                    }
                    match side {
                        0 => borders.top = Some(edge),
                        1 => borders.left = Some(edge),
                        2 => borders.bottom = Some(edge),
                        _ => borders.right = Some(edge),
                    }
                }
                None => match side {
                    0 => top += padding[0],
                    1 => indent_left += padding[1],
                    2 => bottom += padding[2],
                    _ => indent_right += padding[3],
                },
            }
        }
        // style:join-border (default true) merges equal borders of
        // consecutive paragraphs, as DOCX does for identical borders.
        if any_border && borders.top.is_some() && borders.bottom.is_some() {
            borders.between = borders.bottom.clone();
        }
        let autospace = cascade.get(NS_STYLE, "text-autospace").map(|v| v != "none");
        // Justified lines that may compress their spaces: LibreOffice's
        // JustifyLinesWithShrinking, or a minimum word spacing below 100%
        // (`loext:word-spacing-minimum`), as Word's mode-15 justification.
        let justified_compression = self.settings.shrink
            || cascade
                .get(NS_LOEXT, "word-spacing-minimum")
                .and_then(|v| percent(v).ok())
                .is_some_and(|v| v < 1.0);
        // LibreOffice's placement of fixed and minimum lines, within its
        // measured scope: horizontal pages without a layout grid, and
        // paragraphs that do not snap to a page's grid. Vertical pages and
        // snapping paragraphs keep the DOCX placement, which the corpus
        // matches better there.
        let line_extra_above = matches!(
            line_spacing.as_ref().map(|spacing| spacing.rule.as_str()),
            Some("exact" | "atLeast")
        ) && (self.pages.plain_horizontal(&self.current_master)
            || (self.pages.horizontal(&self.current_master)
                && cascade.get(NS_STYLE, "snap-to-layout-grid") == Some("false")));
        let paragraph = DocParagraph {
            alignment: alignment.into(),
            indent_left,
            indent_right,
            indent_first,
            line_spacing,
            keep_next: cascade.get(NS_FO, "keep-with-next") == Some("always"),
            keep_lines: cascade.get(NS_FO, "keep-together") == Some("always"),
            widow_control,
            page_break_before: breaking("break-before") == Some("page")
                && self.story == Story::Body,
            shading,
            borders: any_border.then_some(borders),
            bidi: rtl.then_some(true),
            justified_compression: justified_compression
                && matches!(alignment, "both" | "distribute"),
            // LibreOffice places a minimum line's extra height above the
            // text and a fixed line's baseline at four fifths of its height
            // (SwTextFormatter::CalcRealHeight; rr3 fixed-line controls).
            line_extra_above,
            // LibreOffice applies paragraph indents from the edge of the
            // area beside a wrapped float (measured: a hanging list item
            // beside a left float starts its label at the float's edge plus
            // the first-line position).
            indents_from_floats: true,
            // §20.302 style:punctuation-wrap hanging, §20.296 style:line-break
            // strict, §20.346 style:text-autospace.
            overflow_punct: cascade.get(NS_STYLE, "punctuation-wrap") != Some("simple"),
            kinsoku: cascade.get(NS_STYLE, "line-break").map(|v| v == "strict"),
            auto_space_de: autospace,
            auto_space_dn: autospace,
            snap_to_grid: cascade
                .get(NS_STYLE, "snap-to-layout-grid")
                .map(|v| v == "true"),
            ..DocParagraph::default()
        };
        if breaking("break-before") == Some("column") && self.story == Story::Body {
            self.blocks.push(Block::ColumnBreak);
        }
        Ok(Collected {
            paragraph,
            top,
            bottom,
            contextual: cascade.get(NS_STYLE, "contextual-spacing") == Some("true"),
            named_style: None,
            area: 0,
            page_break_after: breaking("break-after") == Some("page") && self.story == Story::Body,
            column_break_after: breaking("break-after") == Some("column")
                && self.story == Story::Body,
        })
    }

    /// Custom tab stops (§17.8 `style:tab-stops`), positioned from the text
    /// area: LibreOffice measures them from the paragraph indent unless
    /// TabsRelativeToIndent is off (Word-derived documents).
    fn tab_stops(
        &mut self,
        cascade: &Cascade<'a>,
        indent_left: f64,
    ) -> Result<Vec<TabStop>, String> {
        let Some(elements) = cascade.element("tab-stops") else {
            return Ok(Vec::new());
        };
        let origin = if self.settings.tabs_relative.unwrap_or(true) {
            indent_left
        } else {
            0.0
        };
        let mut stops = Vec::new();
        for stop in elements.iter().flat_map(|e| &e.children) {
            let attribute = |ns: &str, name: &str| {
                stop.attributes
                    .get(&(ns.to_string(), name.to_string()))
                    .map(String::as_str)
            };
            let Some(position) = attribute(NS_STYLE, "position").and_then(|v| length_pt(v).ok())
            else {
                continue;
            };
            let alignment = match attribute(NS_STYLE, "type") {
                Some("center") => "center",
                Some("right") => "right",
                Some("char") => "decimal",
                _ => "left",
            };
            let leader_text = attribute(NS_STYLE, "leader-text");
            let leader = match (attribute(NS_STYLE, "leader-style"), leader_text) {
                (_, Some(".")) | (Some("dotted"), _) => "dot",
                (_, Some("-")) | (Some("dash" | "long-dash"), _) => "hyphen",
                (_, Some("_")) | (Some("solid"), _) => "underscore",
                (_, Some("·")) | (Some("dot-dash" | "dot-dot-dash"), _) => "middleDot",
                _ => "none",
            };
            self.charge_construction(std::mem::size_of::<TabStop>() + 16)?;
            stops.push(TabStop {
                pos: origin + position,
                alignment: alignment.into(),
                leader: leader.into(),
            });
        }
        stops.sort_by(|a, b| a.pos.total_cmp(&b.pos));
        Ok(stops)
    }

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

    /// A package file linked from content (Part 2 §3.7, relative to the
    /// package root where content.xml and styles.xml sit): its package path,
    /// which keys the retained resource, and its sniffed media type. Links
    /// outside the package or to missing files yield `None` (nothing to
    /// draw); each file is retained once.
    fn package_image(&mut self, href: &str) -> Result<Option<(String, &'static str)>, String> {
        let path = match iri::resolve("content.xml", href) {
            Some(iri::Target::Package {
                path,
                fragment: None,
            }) if !path.is_empty() && !path.ends_with('/') => path,
            _ => return Ok(None),
        };
        if let Some(mime_type) = self.image_types.get(&path) {
            return Ok(Some((path, mime_type)));
        }
        if !self.package.contains(&path) {
            return Ok(None);
        }
        let bytes = self
            .package
            .session()
            .run_operation("odf-image", |operation| operation.read_bytes(&path))?;
        let Some((bytes, mime_type)) = picture(bytes, self.construction_left()) else {
            return Ok(None);
        };
        self.charge_resource(bytes.len() + 2 * path.len() + 64)?;
        self.image_types.insert(path.clone(), mime_type);
        self.images.push(OwnedDocumentResource {
            key: path.clone(),
            mime_type: mime_type.to_string(),
            bytes,
        });
        Ok(Some((path, mime_type)))
    }

    /// Embedded picture data (`office:binary-data`, base64), retained under a
    /// synthetic key.
    fn binary_image(&mut self, data: &str) -> Result<Option<(String, &'static str)>, String> {
        self.charge_construction(data.len())?;
        let Some(bytes) = image::base64(data) else {
            return Ok(None);
        };
        let Some((bytes, mime_type)) = picture(bytes, self.construction_left()) else {
            return Ok(None);
        };
        // The retained picture (a converted metafile can outgrow its data).
        self.charge_resource(bytes.len() + 64)?;
        let key = format!("odt-binary/{}", self.images.len());
        self.image_types.insert(key.clone(), mime_type);
        self.images.push(OwnedDocumentResource {
            key: key.clone(),
            mime_type: mime_type.to_string(),
            bytes,
        });
        Ok(Some((key, mime_type)))
    }

    /// A package resource's retained bytes: construction bytes that a copy
    /// of a story referencing the resource does not repeat.
    fn charge_resource(&mut self, bytes: usize) -> Result<(), String> {
        self.charge_construction(bytes)?;
        self.resource_bytes += bytes;
        Ok(())
    }

    /// Construction bytes charged so far, less package resources: the
    /// difference across building a story is what each copy of it costs.
    fn copyable_bytes(&self) -> usize {
        self.construction_bytes.saturating_sub(self.resource_bytes)
    }

    /// What the construction budget still admits.
    fn construction_left(&self) -> usize {
        MAX_CONSTRUCTION_BYTES.saturating_sub(self.construction_bytes)
    }

    /// Charges retained model bytes before they are allocated. An overrun is
    /// sticky: the ledger stays exhausted, so a caller that dropped the error
    /// still fails at the next charge and at the end of the read.
    fn charge_construction(&mut self, bytes: usize) -> Result<(), String> {
        match self
            .construction_bytes
            .checked_add(bytes)
            .filter(|total| *total <= MAX_CONSTRUCTION_BYTES)
        {
            Some(total) => {
                self.construction_bytes = total;
                Ok(())
            }
            None => {
                self.construction_bytes = usize::MAX;
                Err("OpenDocument text exceeds the construction limit".to_string())
            }
        }
    }

    /// One paragraph against the document-wide paragraph bound.
    fn count_paragraph(&mut self) -> Result<(), String> {
        self.paragraph_count += 1;
        if self.paragraph_count > MAX_PARAGRAPHS {
            return Err("OpenDocument text exceeds the paragraph limit".into());
        }
        Ok(())
    }

    /// A table cell the reader builds itself (a covered position, a missing
    /// cell, a vertical-merge continuation) with its one empty paragraph:
    /// counted and charged before it is allocated.
    fn charge_empty_cell(&mut self) -> Result<(), String> {
        self.count_paragraph()?;
        self.charge_construction(CELL_BYTES + EMPTY_PARAGRAPH_BYTES)
    }

    /// Fails when a budget was overrun anywhere during the read, even where
    /// the overrun's error was not propagated.
    fn check_budgets(&self) -> Result<(), String> {
        if self.construction_bytes > MAX_CONSTRUCTION_BYTES {
            return Err("OpenDocument text exceeds the construction limit".into());
        }
        if self.cascade_work.get() > MAX_CASCADE_WORK {
            return Err("OpenDocument text exceeds the style resolution limit".into());
        }
        Ok(())
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

    /// Charges style resolution work; an overrun is sticky (see
    /// `charge_construction`).
    fn charge_work(&self, work: usize) -> Result<(), String> {
        match self
            .cascade_work
            .get()
            .checked_add(work)
            .filter(|total| *total <= MAX_CASCADE_WORK)
        {
            Some(total) => {
                self.cascade_work.set(total);
                Ok(())
            }
            None => {
                self.cascade_work.set(usize::MAX);
                Err("OpenDocument text exceeds the style resolution limit".to_string())
            }
        }
    }

    fn occurrence(&mut self, prefix: &str) -> String {
        self.next_occurrence += 1;
        format!("odt/{prefix}/{}", self.next_occurrence)
    }

    /// Paragraph gaps (measured against LibreOffice): the previous block's
    /// bottom margin and the next one's top margin add; `style:contextual-
    /// spacing` on two paragraphs of one named style in one content area
    /// removes the gap (LibreOffice decides by the previous paragraph's
    /// flag). The gap is emitted as the next paragraph's space before with
    /// space after 0, so the DOCX max(after, before) rule reproduces it; a
    /// gap above a table goes to the preceding paragraph's space after.
    fn finish_blocks(&mut self, blocks: Vec<Block>) -> Result<Vec<Finished>, String> {
        let mut out: Vec<Finished> = Vec::with_capacity(blocks.len());
        // The previous flow block's bottom margin and contextual identity.
        let mut previous: Option<(f64, bool, Option<*const Style>, u32)> = None;
        let mut break_pending = false;
        let mut column_pending = false;
        for block in blocks {
            match block {
                Block::Paragraph(collected) => {
                    let mut collected = *collected;
                    let mut paragraph = std::mem::take(&mut collected.paragraph);
                    paragraph.page_break_before |= break_pending;
                    let restart = paragraph.page_break_before || column_pending;
                    paragraph.space_before = match previous {
                        None => collected.top,
                        Some(_) if restart => collected.top,
                        Some((bottom, contextual, named, area)) => {
                            let same = area == collected.area
                                && named.is_some()
                                && named == collected.named_style;
                            if same && contextual {
                                0.0
                            } else {
                                bottom + collected.top
                            }
                        }
                    };
                    // Spacing taller than the page body pushes the paragraph
                    // past any page in LibreOffice too; it is capped so the
                    // paragraph's first line (about 1.5 em) still fits a
                    // page, which the DOCX pagination requires.
                    let line = paragraph.default_font_size.unwrap_or(12.0) * 1.5;
                    paragraph.space_before = paragraph
                        .space_before
                        .min((self.text_height - line).max(0.0));
                    paragraph.space_after = 0.0;
                    paragraph.contextual_spacing = false;
                    break_pending = collected.page_break_after;
                    column_pending = false;
                    previous = Some((
                        collected.bottom,
                        collected.contextual,
                        collected.named_style,
                        collected.area,
                    ));
                    out.push(Finished::Element(BodyElement::Paragraph(Box::new(
                        paragraph,
                    ))));
                    if collected.column_break_after {
                        out.push(Finished::Element(BodyElement::ColumnBreak));
                        column_pending = true;
                    }
                }
                Block::Table(table) => {
                    let TableBlock {
                        table,
                        top,
                        bottom,
                        page_break_before,
                        page_break_after,
                    } = *table;
                    if page_break_before || break_pending {
                        out.push(Finished::Element(BodyElement::PageBreak {
                            origin: None,
                            parity: None,
                            same_paragraph_as_previous: None,
                        }));
                    } else if let Some((bottom_before, ..)) = previous {
                        // The gap above the table: carried by the preceding
                        // paragraph's space after.
                        if let Some(Finished::Element(BodyElement::Paragraph(p))) = out
                            .iter_mut()
                            .rev()
                            .find(|f| !matches!(f, Finished::Section(_)))
                        {
                            p.space_after = bottom_before + top;
                        }
                    }
                    break_pending = page_break_after;
                    column_pending = false;
                    previous = Some((bottom, false, None, self.area));
                    out.push(Finished::Element(BodyElement::Table(Box::new(table))));
                }
                Block::ColumnBreak => {
                    out.push(Finished::Element(BodyElement::ColumnBreak));
                    column_pending = true;
                }
                Block::Section(index) => out.push(Finished::Section(index)),
            }
        }
        if break_pending && self.story == Story::Body {
            out.push(Finished::Element(BodyElement::PageBreak {
                origin: None,
                parity: None,
                same_paragraph_as_previous: None,
            }));
        }
        Ok(out)
    }

    /// A contained story's elements (section marks do not occur there).
    fn finish_story(&mut self, blocks: Vec<Block>) -> Result<Vec<BodyElement>, String> {
        Ok(self
            .finish_blocks(blocks)?
            .into_iter()
            .filter_map(|finished| match finished {
                Finished::Element(element) => Some(element),
                Finished::Section(_) => None,
            })
            .collect())
    }

    /// Cell content: paragraphs and nested tables.
    fn finish_cell(&mut self, elements: Vec<BodyElement>) -> Vec<CellElement> {
        elements
            .into_iter()
            .filter_map(|element| match element {
                BodyElement::Paragraph(p) => Some(CellElement::Paragraph(p)),
                BodyElement::Table(t) => Some(CellElement::Table(t)),
                _ => None,
            })
            .collect()
    }

    fn finish_body(
        &mut self,
        blocks: Vec<Block>,
    ) -> Result<(Vec<BodyElement>, SectionMarks), String> {
        let finished = self.finish_blocks(blocks)?;
        let mut body = Vec::with_capacity(finished.len());
        // (body index before which a section starts, section index)
        let mut marks = Vec::new();
        for item in finished {
            match item {
                Finished::Element(element) => body.push(element),
                Finished::Section(index) => marks.push((body.len(), index)),
            }
        }
        if !self.pending_anchors.is_empty() {
            let runs = std::mem::take(&mut self.pending_anchors);
            body.push(BodyElement::Paragraph(Box::new(anchor_paragraph(runs))));
        }
        Ok((body, marks))
    }
}

/// (body index before which a section starts, section index).
type SectionMarks = Vec<(usize, usize)>;

enum Finished {
    Element(BodyElement),
    Section(usize),
}

/// An empty paragraph carrying drawings left without a following paragraph.
fn anchor_paragraph(runs: Vec<DocRun>) -> DocParagraph {
    let host = TextRun {
        font_size: 12.0,
        ..TextRun::default()
    };
    DocParagraph {
        alignment: "left".into(),
        runs: runs
            .into_iter()
            .flat_map(|run| text::with_anchor_host(run, &host))
            .collect(),
        ..DocParagraph::default()
    }
}

/// The text-properties cascade of a paragraph chain.
fn cascade_text<'a>(chain: &[StyleRef<'a>]) -> Cascade<'a> {
    Cascade::new(chain.to_vec(), TEXT)
}

/// A relative length (`NN%`) resolved against the nearest absolute value of
/// the same attribute further up the chain (§20.204: percentages refer to
/// the parent style's value).
fn relative_length(
    chain: &[StyleRef<'_>],
    group: Group,
    namespace: &str,
    name: &str,
    relative: &str,
) -> f64 {
    let key = (namespace.to_string(), name.to_string());
    let mut factor = percent(relative).unwrap_or(1.0);
    let mut seen_first = false;
    for style in chain {
        let Some(value) = style.group(group).and_then(|g| g.attributes.get(&key)) else {
            continue;
        };
        if !seen_first {
            seen_first = true;
            continue;
        }
        if value.ends_with('%') {
            factor *= percent(value).unwrap_or(1.0);
            continue;
        }
        return length_pt(value).unwrap_or(0.0) * factor;
    }
    0.0
}

/// A paragraph's background (§20.175 fo:background-color, or LibreOffice's
/// area fill `loext:graphic-properties`, whose solid colour wins when both
/// are written; a gradient or bitmap fill draws its first colour).
fn paragraph_fill(paragraph: &Cascade<'_>, fill: &Cascade<'_>) -> Option<String> {
    match fill.get(NS_DRAW, "fill") {
        Some("none") => return None,
        Some("solid") => {
            if let Some(color) = fill
                .get(NS_DRAW, "fill-color")
                .and_then(|c| hex_color(c).ok())
            {
                return Some(color);
            }
        }
        Some("gradient" | "hatch" | "bitmap") => {
            if let Some(color) = fill
                .get(NS_DRAW, "fill-color")
                .and_then(|c| hex_color(c).ok())
            {
                return Some(color);
            }
        }
        _ => {}
    }
    match paragraph.get(NS_FO, "background-color") {
        None | Some("transparent") => None,
        Some(value) => hex_color(value).ok(),
    }
}

/// The four sides' borders (top, left, bottom, right) of a property group:
/// `fo:border` and its side-specific forms (§20.183), side wins.
fn border_edges(cascade: &Cascade<'_>) -> [Option<ParaBorderEdge>; 4] {
    let all = cascade.get(NS_FO, "border");
    let widths = cascade.get(NS_STYLE, "border-line-width");
    ["top", "left", "bottom", "right"].map(|side| {
        let value = cascade.get(NS_FO, &format!("border-{side}")).or(all)?;
        let width = cascade
            .get(NS_STYLE, &format!("border-line-width-{side}"))
            .or(widths);
        let spec = parse_border(value, width)?;
        Some(ParaBorderEdge {
            style: spec.style.to_string(),
            color: spec.color,
            width: spec.width,
            space: 0.0,
        })
    })
}

/// `fo:padding` and its sides (§20.217), side wins.
fn padding_values(cascade: &Cascade<'_>) -> [f64; 4] {
    let all = cascade.get(NS_FO, "padding");
    ["top", "left", "bottom", "right"].map(|side| {
        cascade
            .get(NS_FO, &format!("padding-{side}"))
            .or(all)
            .and_then(|v| length_pt(v).ok())
            .unwrap_or(0.0)
            .max(0.0)
    })
}

/// A parsed border: DOCX style name, colour and width (pt).
pub(crate) struct BorderSpecOdf {
    pub style: &'static str,
    pub color: Option<String>,
    pub width: f64,
}

/// `fo:border` (§20.183, XSL `border`): "none", or a width, a style and a
/// colour in any order. A double line's total width is the inner, gap and
/// outer widths of `style:border-line-width` (§20.185) when given.
pub(crate) fn parse_border(value: &str, line_width: Option<&str>) -> Option<BorderSpecOdf> {
    let value = value.trim();
    if value.is_empty() || value == "none" || value == "hidden" {
        return None;
    }
    let mut width = None;
    let mut style = None;
    let mut color = None;
    for part in value.split_whitespace() {
        if let Ok(length) = length_pt(part) {
            width = Some(length);
        } else if part.starts_with('#') {
            color = hex_color(part).ok();
        } else {
            match part {
                "none" | "hidden" => return None,
                "thin" => width = Some(0.75),
                "medium" => width = Some(2.25),
                "thick" => width = Some(4.5),
                "solid" | "double" | "dotted" | "dashed" | "groove" | "ridge" | "inset"
                | "outset" | "dash-dot" | "dash-dot-dot" | "fine-dashed" | "double-thin" => {
                    style = Some(part)
                }
                _ => {}
            }
        }
    }
    let mut width = width.unwrap_or(0.75);
    if let Some(parts) = line_width {
        let sum: f64 = parts
            .split_whitespace()
            .filter_map(|p| length_pt(p).ok())
            .sum();
        if sum > 0.0 {
            width = sum;
        }
    }
    if width <= 0.0 {
        return None;
    }
    let style = match style.unwrap_or("solid") {
        "double" | "double-thin" => "double",
        "dotted" => "dotted",
        "dashed" | "fine-dashed" => "dashed",
        "dash-dot" => "dotDash",
        "dash-dot-dot" => "dotDotDash",
        "groove" => "threeDEngrave",
        "ridge" => "threeDEmboss",
        "inset" => "inset",
        "outset" => "outset",
        _ => "single",
    };
    Some(BorderSpecOdf {
        style,
        color,
        width,
    })
}

/// ODF lengths (Part 3 §18.3.18): cm, mm, in, pt, pc, px (1/96 in), and the
/// LibreOffice twip unit "twip"? (not ODF; not accepted).
pub fn length_pt(value: &str) -> Result<f64, String> {
    let value = value.trim();
    let split = value
        .find(|c: char| c.is_ascii_alphabetic())
        .ok_or_else(|| invalid("a length has no unit"))?;
    let (number, unit) = value.split_at(split);
    let number: f64 = number
        .trim()
        .parse()
        .map_err(|_| invalid("a length is not a number"))?;
    let factor = match unit {
        "pt" => 1.0,
        "cm" => 72.0 / 2.54,
        "mm" => 72.0 / 25.4,
        "in" | "inch" => 72.0,
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
        .trim()
        .parse()
        .map_err(|_| invalid("a percentage is not a number"))?;
    if !number.is_finite() || !(-100_000.0..=100_000.0).contains(&number) {
        return Err(invalid("a percentage is out of range"));
    }
    Ok(number / 100.0)
}

fn hex_color(value: &str) -> Result<String, String> {
    let hex = value
        .trim()
        .strip_prefix('#')
        .filter(|h| h.len() == 6 && h.chars().all(|c| c.is_ascii_hexdigit()))
        .ok_or_else(|| invalid("a colour is not #rrggbb"))?;
    Ok(hex.to_ascii_lowercase())
}

/// Facts of a run's font for an object placed in its line.
fn font_facts_of(run: &TextRun) -> RunFontFacts {
    RunFontFacts {
        font_family: run.font_family.clone(),
        font_family_high_ansi: run.font_family_high_ansi.clone(),
        font_family_east_asia: run.font_family_east_asia.clone(),
        font_family_cs: run.font_family_cs.clone(),
        font_size: Some(run.font_size),
        font_size_cs: run.font_size_cs,
        bold: run.bold,
        italic: run.italic,
        ..RunFontFacts::default()
    }
}

#[cfg(test)]
#[path = "odt_tests.rs"]
mod tests;

/// Picture bytes in a format the renderer draws, with their media type. A
/// StarView metafile (LibreOffice's object replacement format) is converted
/// to SVG.
/// A picture's retained bytes and media type; a StarView metafile converts to
/// SVG within `budget` bytes.
fn picture(bytes: Vec<u8>, budget: usize) -> Option<(Vec<u8>, &'static str)> {
    if let Some(mime_type) = image::sniff(&bytes) {
        return Some((bytes, mime_type));
    }
    crate::svm::to_svg(&bytes, budget).map(|converted| (converted.svg, "image/svg+xml"))
}
