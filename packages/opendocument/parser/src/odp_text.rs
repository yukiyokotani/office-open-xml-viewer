//! Text of drawing objects (ODF 1.3 Part 3 §5 paragraphs, §5.3 lists,
//! §6 fields; §20 text/paragraph properties) mapped onto the PPTX text
//! model.
//!
//! Formatting cascades per attribute (§16.2) through, nearest first: the
//! enclosing spans' styles, the paragraph style, the object's
//! `draw:text-style-name`, the object's graphic or presentation style chain
//! (for outline placeholders the per-level `<master>-outline<n>` styles, as
//! LibreOffice applies them) and the default graphic style.

use crate::odp_media::Media;
use crate::odp_style::{Nd, Props, Scope, Sheet, GRAPHIC, PARAGRAPH, TEXT};
use crate::odp_units::{color, is_dark, length_emu, percent, with_alpha, EMU_PER_PT};
use crate::xml::{self, NS_DRAW, NS_FO, NS_LOEXT, NS_STYLE, NS_TEXT, NS_XLINK};
use ooxml_common::effect::Shadow;
use ooxml_common::text::SpaceLine;
use pptx_model::{Bullet, Paragraph, TabStop, TextBody, TextOutline, TextRun, TextRunData};
use std::cell::Cell;
use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;

/// Implementation bound on the text work of the whole deck (every object,
/// master objects once per slide that shows them).
pub const MAX_DECK_TEXT_WORK: usize = 256 * 1024 * 1024;

/// The remaining text payload and work of one object (see
/// [`MAX_OBJECT_TEXT_BYTES`], [`MAX_OBJECT_TEXT_WORK`]); work is also
/// charged to the deck's [`MAX_DECK_TEXT_WORK`].
pub struct TextPayload<'d> {
    bytes: Cell<usize>,
    work: Cell<usize>,
    deck_work: &'d Cell<usize>,
}

impl<'d> TextPayload<'d> {
    pub fn new(deck_work: &'d Cell<usize>) -> Self {
        Self {
            bytes: Cell::new(MAX_OBJECT_TEXT_BYTES),
            work: Cell::new(MAX_OBJECT_TEXT_WORK),
            deck_work,
        }
    }
    /// Charge `bytes` before they are retained.
    pub fn charge(&self, bytes: usize) -> Result<(), String> {
        let left = self
            .bytes
            .get()
            .checked_sub(bytes)
            .ok_or_else(|| "OpenDocument text exceeds the text limit".to_string())?;
        self.bytes.set(left);
        Ok(())
    }
    /// Charge `steps` of text work (nodes and style records visited).
    pub fn step(&self, steps: usize) -> Result<(), String> {
        let error = || "OpenDocument text exceeds the text work limit".to_string();
        let left = self.work.get().checked_sub(steps).ok_or_else(error)?;
        let deck_left = self.deck_work.get().checked_sub(steps).ok_or_else(error)?;
        self.work.set(left);
        self.deck_work.set(deck_left);
        Ok(())
    }
}

/// Retained bytes of a model record: its fixed size plus its serialized
/// strings and lists (the JSON size bounds what the record owns on the heap
/// beyond its fixed part).
pub fn record_bytes<T: serde::Serialize>(value: &T) -> Result<usize, String> {
    let measured = ooxml_common::json_measurement::measure_json(value)?;
    usize::try_from(measured.json_bytes)
        .ok()
        .and_then(|n| n.checked_add(std::mem::size_of::<T>()))
        .ok_or_else(|| "OpenDocument object exceeds the retained-size limit".to_string())
}

pub const NS_PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

/// Implementation bound on paragraphs per text body.
const MAX_PARAGRAPHS: usize = 100_000;
/// Implementation bound on the text payload of one object: a shape's text
/// body, or every cell of one table together. It counts the bytes each run
/// and paragraph retains (text, strings, and the fixed size of the model
/// record), charged before the record is kept.
pub const MAX_OBJECT_TEXT_BYTES: usize = 32 * 1024 * 1024;
/// Implementation bound on the inline nodes one object's text visits,
/// weighted by the span nesting each one resolves its style through.
pub const MAX_OBJECT_TEXT_WORK: usize = 16 * 1024 * 1024;
/// Implementation bound on the text of one slide's speaker notes (bytes).
const MAX_NOTES_BYTES: usize = 16 * 1024 * 1024;
/// Implementation bound on element nesting inside a paragraph.
const MAX_SPAN_DEPTH: usize = 64;

/// The styles an object's text inherits from the object itself.
#[derive(Clone, Default)]
pub struct ObjectStyles<'a> {
    /// The object's style and its parents (graphic or presentation family).
    pub chain: Vec<Nd<'a>>,
    /// The object's automatic style alone (outline levels keep its own
    /// attributes above the level style).
    pub automatic: Option<Nd<'a>>,
    /// Master page whose `-outline<n>` styles apply, for outline objects.
    pub outline_master: Option<&'a str>,
    /// `draw:text-style-name` (paragraph family) and its parents.
    pub text_style: Vec<Nd<'a>>,
    pub defaults: Vec<Nd<'a>>,
}

impl<'a> ObjectStyles<'a> {
    /// Style nodes for a paragraph at list depth `level` (0-based).
    pub fn for_level(&self, sheet: &Sheet<'a>, level: usize) -> Vec<Nd<'a>> {
        let mut out = self.text_style.clone();
        match self.outline_master {
            Some(master) if level > 0 => {
                if let Some(auto) = self.automatic {
                    out.push(auto);
                }
                let name = format!("{master}-outline{}", level + 1);
                let mut current = sheet.common("presentation", &name);
                let mut guard = 0;
                while let Some(style) = current {
                    out.push(style);
                    guard += 1;
                    if guard > 64 {
                        break;
                    }
                    current = style
                        .attribute((NS_STYLE, "parent-style-name"))
                        .and_then(|p| sheet.common("presentation", p));
                }
                if guard == 0 {
                    out.extend(self.chain.iter().copied());
                }
            }
            _ => out.extend(self.chain.iter().copied()),
        }
        out.extend(self.defaults.iter().copied());
        out
    }
}

pub struct TextEnv<'a, 'e> {
    pub sheet: &'a Sheet<'a>,
    pub scope: Scope,
    pub base: &'a str,
    pub object: ObjectStyles<'a>,
    pub slide_number: usize,
    pub slide_count: usize,
    /// Texts of the slide's `presentation:footer`, `header` and `date-time`.
    /// Shared with every object of the slide (one copy per declaration).
    pub footer: Option<Rc<str>>,
    pub header: Option<Rc<str>>,
    pub date_time: Option<Rc<str>>,
    /// Background behind the text, for automatic font colours.
    pub background: Option<String>,
    /// Slide names (`draw:name`) to slide indexes, for internal links.
    pub pages: &'e HashMap<String, usize>,
    /// The object's text payload (shared by every cell of a table).
    pub payload: &'e TextPayload<'e>,
}

/// Paragraph context while walking lists.
#[derive(Clone)]
struct ListContext<'a> {
    /// Depth (1-based) of the enclosing `text:list`, 0 outside lists.
    depth: usize,
    /// The list style in effect.
    style: Option<Nd<'a>>,
    /// Whether the paragraph is the first one of a `text:list-item` (labels
    /// apply to it; headers and following paragraphs have none).
    labelled: bool,
    start_value: Option<u32>,
    /// The list instance the paragraph numbers in (§5.3.1: a top-level
    /// `text:list` starts one unless it continues an earlier list); `None`
    /// outside lists.
    instance: Option<usize>,
    /// `xml:id` of the innermost `text:list`, which a later list may name in
    /// `text:continue-list`.
    list_id: Option<&'a str>,
    /// How a nested list's first item continues (`text:continue-list`,
    /// `text:continue-numbering`); applies to `first_in_list` only.
    resume: Resume<'a>,
    /// Whether the paragraph is the first labelled one of its `text:list`.
    first_in_list: bool,
    /// The numbering style in effect for continuation checks (§19.786):
    /// the list's own style, else the one its first paragraph's styles
    /// select, else the enclosing list's.
    numbering_style: Option<Nd<'a>>,
}

/// The continuation a nested `text:list` asks for (§19.785, §19.786).
#[derive(Clone, Copy, Default, PartialEq)]
enum Resume<'a> {
    #[default]
    None,
    /// `text:continue-list`: the list with this `xml:id`.
    List(&'a str),
    /// `text:continue-numbering="true"`: the preceding list at this level
    /// when it has the same list style.
    Preceding,
    /// `text:continue-numbering="false"`: a new sequence.
    Restart,
    /// `text:continue-list` naming a list outside this text: continue
    /// after that list's last number.
    After(u32),
}

impl ListContext<'_> {
    /// The context of a paragraph outside lists.
    fn outside() -> Self {
        ListContext {
            depth: 0,
            style: None,
            labelled: false,
            start_value: None,
            instance: None,
            list_id: None,
            resume: Resume::None,
            first_in_list: false,
            numbering_style: None,
        }
    }
}

/// List instances of one text body, for `text:continue-numbering` and
/// `text:continue-list` (§19.785, §19.786).
struct ListInstances<'a, 'p> {
    /// Charged for every block and list node visited.
    payload: &'p TextPayload<'p>,
    sheet: &'a Sheet<'a>,
    scope: Scope,
    /// The styles of the object the text belongs to (the list style
    /// cascade continues through them).
    object: &'p ObjectStyles<'a>,
    count: usize,
    /// The last top-level list: its instance and list style name.
    previous: Option<(usize, Option<Nd<'a>>)>,
    /// `xml:id` of top-level lists to their instance.
    ids: HashMap<&'a str, usize>,
}

impl<'a> ListInstances<'a, '_> {
    /// The instance a top-level `text:list` numbers in: the list it names
    /// by `text:continue-list`; with `text:continue-numbering="true"` the
    /// preceding list when it has the same list style (§19.786); with
    /// `"false"` a new one. §19.786 defines no value for an absent
    /// attribute: the list then continues the preceding list of the text,
    /// as LibreOffice numbers consecutive lists (its own files mark only
    /// some continuations).
    fn top_level(
        &mut self,
        list: Nd<'a>,
        style: Option<Nd<'a>>,
        depth: usize,
    ) -> Result<(usize, Resume<'a>), String> {
        // A continued list outside this text (another frame or slide) is
        // continued after its own last number.
        let resume = match list.attribute((NS_TEXT, "continue-list")) {
            Some(id) if !self.ids.contains_key(id) => self
                .external_last(id, depth)?
                .map_or(Resume::None, Resume::After),
            // A list of this text: after that list's own last number.
            Some(id) => Resume::List(id),
            None => Resume::None,
        };
        let continued = match list.attribute((NS_TEXT, "continue-list")) {
            Some(id) => self.ids.get(id).copied(),
            None => match list.attribute((NS_TEXT, "continue-numbering")) {
                Some("false") => None,
                Some("true") => self
                    .previous
                    .filter(|(_, previous_style)| *previous_style == style)
                    .map(|(instance, _)| instance),
                _ => self.previous.map(|(instance, _)| instance),
            },
        };
        let instance = continued.unwrap_or_else(|| {
            self.count += 1;
            self.count
        });
        if let Some(id) = list.attribute((crate::xml::NS_XML, "id")) {
            self.ids.insert(id, instance);
        }
        self.previous = Some((instance, style));
        Ok((instance, resume))
    }

    /// The last item number of the list with `xml:id` `id` anywhere in the
    /// document: the numbering of the text that holds it, resolved once per
    /// text for the whole deck with the same list semantics as a text body
    /// (every labelled item counts as numbered). Nested resolutions are
    /// bounded; a cycle resolves to no number.
    /// `depth` is the list nesting of the caller: the resolution continues
    /// it, so one nesting bound covers rendering and resolution together.
    fn external_last(&self, id: &str, depth: usize) -> Result<Option<u32>, String> {
        let Some(&(scope, list)) = self.sheet.list_ids.get(id) else {
            return Ok(None);
        };
        let container = text_container(list);
        let numbers = list_numbers(self.sheet, container, scope, self.payload, depth + 1)?;
        Ok(numbers.and_then(|n| n.get(id).copied()))
    }
}

/// The element whose text holds `list`: its nearest ancestor that is not
/// part of the list structure.
fn text_container<'a>(list: Nd<'a>) -> Nd<'a> {
    let mut node = list;
    while let Some(parent) = node.parent_element() {
        let structural = xml::is(parent, NS_TEXT, "list")
            || xml::is(parent, NS_TEXT, "list-item")
            || xml::is(parent, NS_TEXT, "list-header")
            || xml::is(parent, NS_TEXT, "section");
        if !structural {
            return parent;
        }
        node = parent;
    }
    node
}

/// The last number of each identified list of one text, cached in the
/// sheet for the deck (`None` while it is being resolved, or past the
/// nesting bound).
fn list_numbers<'a>(
    sheet: &'a Sheet<'a>,
    container: Nd<'a>,
    scope: crate::odp_style::Scope,
    payload: &TextPayload<'_>,
    nesting: usize,
) -> Result<Option<Rc<HashMap<String, u32>>>, String> {
    let key = (
        container.document() as *const roxmltree::Document<'_> as usize,
        container.id(),
    );
    if let Some(cached) = sheet.list_numbers.borrow().get(&key) {
        payload.step(1)?;
        return Ok(cached.clone());
    }
    if nesting > MAX_SPAN_DEPTH {
        return Ok(None);
    }
    sheet.list_numbers.borrow_mut().insert(key, None);
    let object = container_object_styles(sheet, scope, container, payload)?;
    let mut instances = ListInstances {
        payload,
        sheet,
        scope,
        object: &object,
        count: 0,
        previous: None,
        ids: HashMap::new(),
    };
    let mut items: Vec<(Nd<'a>, ListContext<'a>)> = Vec::new();
    collect_blocks(
        container,
        ListContext::outside(),
        &mut items,
        &mut instances,
        nesting,
    )?;
    let mut numbering = Numbering::default();
    for (node, list) in items.iter().filter(|(_, l)| l.labelled && l.depth > 0) {
        // The list style as rendering finds it: the list's, else one its
        // paragraph styles select (an object's own list style is not
        // known here).
        let style = list.numbering_style.or_else(|| {
            paragraph_chain(sheet, scope, *node)
                .iter()
                .find_map(|s| s.attribute((NS_STYLE, "list-style-name")))
                .and_then(|name| sheet.list(scope, name))
        });
        let level = style
            .map(|style| list_level(style, list.depth, payload))
            .transpose()?
            .flatten();
        // As rendering counts: only a numbered level takes a number; a
        // bullet, picture or missing level records an unnumbered item.
        let Some(level) = level.filter(|l| xml::is(*l, NS_TEXT, "list-level-style-number")) else {
            numbering.record(list, None);
            continue;
        };
        let counted = consecutive_context(level, list);
        let start = level
            .attribute((NS_TEXT, "start-value"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1);
        let number = match (counted.start_value, numbering.previous(&counted)) {
            (Some(value), _) => value,
            (None, Some(n)) => n.saturating_add(1),
            (None, None) => start,
        };
        numbering.record(&counted, Some(number));
    }
    let numbers: Rc<HashMap<String, u32>> = Rc::new(
        numbering
            .last_of_list
            .into_iter()
            .map(|(id, n)| (id.to_string(), n))
            .collect(),
    );
    sheet
        .list_numbers
        .borrow_mut()
        .insert(key, Some(numbers.clone()));
    Ok(Some(numbers))
}

/// Build the text body of an object from its text container (the
/// `draw:text-box` or the shape element itself). `None` when the object has
/// no paragraphs.
pub fn text_body<'a>(
    container: Nd<'a>,
    env: &TextEnv<'a, '_>,
    media: &mut Media<'_>,
    graphic: &Props<'a>,
) -> Result<Option<TextBody>, String> {
    let mut items: Vec<(Nd<'a>, ListContext<'a>)> = Vec::new();
    let mut instances = ListInstances {
        payload: env.payload,
        sheet: env.sheet,
        scope: env.scope,
        object: &env.object,
        count: 0,
        previous: None,
        ids: HashMap::new(),
    };
    collect_blocks(
        container,
        ListContext::outside(),
        &mut items,
        &mut instances,
        0,
    )?;
    if items.is_empty() {
        return Ok(None);
    }
    let mut body = body_frame(graphic);
    let mut counters = Numbering::default();
    for (node, list) in &items {
        let paragraph = paragraph(*node, list, env, media, &mut counters)?;
        body.paragraphs.push(paragraph);
    }
    Ok(Some(body))
}

fn collect_blocks<'a>(
    parent: Nd<'a>,
    list: ListContext<'a>,
    out: &mut Vec<(Nd<'a>, ListContext<'a>)>,
    instances: &mut ListInstances<'a, '_>,
    depth: usize,
) -> Result<(), String> {
    if depth > MAX_SPAN_DEPTH {
        return Err(crate::error::invalid("lists are nested too deeply"));
    }
    for child in parent.children() {
        instances.payload.step(1)?;
        if !child.is_element() {
            continue;
        }
        if xml::is(child, NS_TEXT, "p") || xml::is(child, NS_TEXT, "h") {
            push_block(out, (child, list.clone()), instances.payload)?;
        } else if xml::is(child, NS_TEXT, "list") {
            collect_list(child, &list, out, instances, depth + 1)?;
        } else if xml::is(child, NS_TEXT, "section") || xml::is(child, NS_TEXT, "soft-page-break") {
            collect_blocks(child, list.clone(), out, instances, depth + 1)?;
        }
    }
    Ok(())
}

/// Keep one collected block, charged to the object's text payload (shared
/// by every resolution in progress) before it is kept.
fn push_block<'a>(
    out: &mut Vec<(Nd<'a>, ListContext<'a>)>,
    item: (Nd<'a>, ListContext<'a>),
    payload: &TextPayload<'_>,
) -> Result<(), String> {
    if out.len() >= MAX_PARAGRAPHS {
        return Err("OpenDocument text exceeds the paragraph limit".to_string());
    }
    payload.charge(std::mem::size_of::<(Nd<'a>, ListContext<'a>)>())?;
    out.push(item);
    Ok(())
}

/// A nested `text:list` inside a list item: one level deeper, inheriting the
/// enclosing list style unless it names its own.
fn collect_list<'a>(
    list_node: Nd<'a>,
    outer: &ListContext<'a>,
    out: &mut Vec<(Nd<'a>, ListContext<'a>)>,
    instances: &mut ListInstances<'a, '_>,
    depth: usize,
) -> Result<(), String> {
    if depth > MAX_SPAN_DEPTH {
        return Err(crate::error::invalid("lists are nested too deeply"));
    }
    let style = if list_node.attribute((NS_TEXT, "style-name")).is_some() {
        Some(list_node)
    } else {
        outer.style
    };
    let numbering_style = effective_list_style(instances, list_node, outer)?;
    // A nested list numbers within its enclosing list's instance, at its
    // own level; an explicit continuation resumes that level's sequence.
    let list_id = list_node.attribute((crate::xml::NS_XML, "id"));
    let (instance, resume) = match outer.instance {
        Some(instance) => {
            if let Some(id) = list_id {
                instances.ids.insert(id, instance);
            }
            let resume = match (
                list_node.attribute((NS_TEXT, "continue-list")),
                list_node.attribute((NS_TEXT, "continue-numbering")),
            ) {
                (Some(id), _) if !instances.ids.contains_key(id) => instances
                    .external_last(id, depth)?
                    .map_or(Resume::Restart, Resume::After),
                (Some(id), _) => Resume::List(id),
                (None, Some("true")) => Resume::Preceding,
                (None, Some("false")) => Resume::Restart,
                // Without either attribute a nested list continues its
                // level within the enclosing list, as LibreOffice numbers
                // the per-paragraph lists it writes.
                (None, _) => Resume::None,
            };
            (instance, resume)
        }
        None => instances.top_level(list_node, numbering_style, depth)?,
    };
    let inner = ListContext {
        depth: outer.depth + 1,
        style,
        labelled: false,
        start_value: None,
        instance: Some(instance),
        list_id,
        resume,
        first_in_list: false,
        numbering_style,
    };
    let mut first_labelled = true;
    for item in list_node.children() {
        instances.payload.step(1)?;
        if !item.is_element() {
            continue;
        }
        let header = xml::is(item, NS_TEXT, "list-header");
        if !(header || xml::is(item, NS_TEXT, "list-item")) {
            continue;
        }
        let mut first = !header;
        let start_value = item
            .attribute((NS_TEXT, "start-value"))
            .and_then(|v| v.parse().ok());
        // §19.881 text:style-override: the item (and its sub-lists) use
        // that list style instead of the list's.
        let overridden;
        let inner = match item
            .attribute((NS_TEXT, "style-override"))
            .and_then(|name| instances.sheet.list(instances.scope, name))
        {
            Some(style) => {
                overridden = ListContext {
                    style: Some(item),
                    numbering_style: Some(style),
                    ..inner.clone()
                };
                &overridden
            }
            None => &inner,
        };
        for block in item.children() {
            instances.payload.step(1)?;
            if !block.is_element() {
                continue;
            }
            if xml::is(block, NS_TEXT, "list") {
                collect_list(block, inner, out, instances, depth + 1)?;
                first = false;
            } else if xml::is(block, NS_TEXT, "p") || xml::is(block, NS_TEXT, "h") {
                push_block(
                    out,
                    (
                        block,
                        ListContext {
                            labelled: first,
                            start_value: if first { start_value } else { None },
                            first_in_list: first && first_labelled,
                            ..inner.clone()
                        },
                    ),
                    instances.payload,
                )?;
                if first {
                    first_labelled = false;
                }
                first = false;
            }
        }
    }
    Ok(())
}

/// The text frame properties of an object (§20: padding, anchoring,
/// wrapping, autofit, columns, writing mode).
pub fn body_frame(graphic: &Props<'_>) -> TextBody {
    let inset = |name: &str| {
        graphic
            .side(NS_FO, name, "padding")
            .and_then(length_emu)
            .unwrap_or(0.0)
            .round() as i64
    };
    let anchor = match graphic.get(NS_DRAW, "textarea-vertical-align") {
        Some("middle") => "ctr",
        Some("bottom") => "b",
        _ => "t",
    };
    let wrap = if graphic.get(NS_FO, "wrap-option") == Some("no-wrap") {
        "none"
    } else {
        "square"
    };
    let shrink = graphic.get(NS_STYLE, "shrink-to-fit") == Some("true")
        || graphic.get(NS_DRAW, "fit-to-size") == Some("shrink-to-fit");
    let auto_fit = if shrink {
        "norm"
    } else if graphic.get(NS_DRAW, "auto-grow-height") == Some("true") {
        "sp"
    } else {
        "none"
    };
    let vert = match graphic
        .get(NS_STYLE, "writing-mode")
        .or_else(|| graphic.get(NS_LOEXT, "writing-mode"))
    {
        Some("tb-rl" | "tb") => "eaVert",
        Some("bt-lr") => "vert270",
        Some("tb-lr") => "mongolianVert",
        _ => "horz",
    };
    let (num_col, spc_col) = match graphic.child(NS_STYLE, "columns") {
        Some(columns) => (
            columns
                .attribute((NS_FO, "column-count"))
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(1)
                .clamp(1, 16),
            columns
                .attribute((NS_FO, "column-gap"))
                .and_then(length_emu)
                .unwrap_or(0.0)
                .round() as i64,
        ),
        None => (1, 0),
    };
    TextBody {
        vertical_anchor: anchor.to_string(),
        paragraphs: Vec::new(),
        default_font_size: None,
        default_bold: None,
        default_italic: None,
        l_ins: inset("padding-left"),
        r_ins: inset("padding-right"),
        t_ins: inset("padding-top"),
        b_ins: inset("padding-bottom"),
        wrap: wrap.to_string(),
        vert: vert.to_string(),
        auto_fit: auto_fit.to_string(),
        font_scale: None,
        ln_spc_reduction: None,
        num_col,
        spc_col,
        rtl_col: false,
        spc_first_last_para: false,
        anchor_ctr: graphic.get(NS_DRAW, "textarea-horizontal-align") == Some("center"),
        compat_ln_spc: None,
        text_warp: None,
        shrink_on_overflow: shrink,
    }
}

/// The list style node a list context names: the nearest `text:list` with a
/// `text:style-name`, else the object's `text:list-style` / list style name.
fn list_style<'a>(
    list: &ListContext<'a>,
    env: &TextEnv<'a, '_>,
    styles: &[Nd<'a>],
) -> Option<Nd<'a>> {
    resolve_list_style(env.sheet, env.scope, list.style, styles)
}

/// The list style cascade (shared by rendering and continuation checks):
/// the nearest `text:list` naming a style, else the first of `styles` that
/// selects one or embeds one in its graphic properties.
fn resolve_list_style<'a>(
    sheet: &Sheet<'a>,
    scope: Scope,
    named: Option<Nd<'a>>,
    styles: &[Nd<'a>],
) -> Option<Nd<'a>> {
    // A list names its style; an overriding list item names its own.
    if let Some(name) = named.and_then(|l| {
        l.attribute((NS_TEXT, "style-name"))
            .or_else(|| l.attribute((NS_TEXT, "style-override")))
    }) {
        if let Some(found) = sheet.list(scope, name) {
            return Some(found);
        }
    }
    for style in styles {
        if let Some(name) = style.attribute((NS_STYLE, "list-style-name")) {
            if let Some(found) = sheet.list(scope, name) {
                return Some(found);
            }
        }
        if let Some(graphic) = crate::odp_style::property_element(*style, GRAPHIC) {
            if let Some(embedded) = graphic
                .children()
                .find(|c| xml::is(*c, NS_TEXT, "list-style"))
            {
                return Some(embedded);
            }
        }
    }
    None
}

/// §19.784 `text:consecutive-numbering="true"`: the list style numbers all
/// levels in one sequence, counted as one level.
fn consecutive_context<'a>(level_node: Nd<'a>, list: &ListContext<'a>) -> ListContext<'a> {
    let consecutive = level_node
        .parent_element()
        .and_then(|style| style.attribute((NS_TEXT, "consecutive-numbering")))
        == Some("true");
    if consecutive {
        ListContext {
            depth: 1,
            ..list.clone()
        }
    } else {
        list.clone()
    }
}

/// The level style of `level`, each record inspected charged as work.
/// §16.32: a level the list style does not define uses the next lower
/// defined level.
fn list_level<'a>(
    list_style: Nd<'a>,
    level: usize,
    payload: &TextPayload<'_>,
) -> Result<Option<Nd<'a>>, String> {
    let mut lower: Option<(usize, Nd<'a>)> = None;
    for n in list_style.children() {
        payload.step(1)?;
        let Some(defined) = n
            .attribute((NS_TEXT, "level"))
            .and_then(|v| v.parse::<usize>().ok())
        else {
            continue;
        };
        if defined == level {
            return Ok(Some(n));
        }
        if defined < level && lower.is_none_or(|(best, _)| defined > best) {
            lower = Some((defined, n));
        }
    }
    Ok(lower.map(|(_, n)| n))
}

fn paragraph_chain<'a>(sheet: &Sheet<'a>, scope: Scope, node: Nd<'a>) -> Vec<Nd<'a>> {
    node.attribute((NS_TEXT, "style-name"))
        .and_then(|name| sheet.chain(scope, "paragraph", name).ok())
        .unwrap_or_default()
}

fn paragraph<'a>(
    node: Nd<'a>,
    list: &ListContext<'a>,
    env: &TextEnv<'a, '_>,
    media: &mut Media<'_>,
    counters: &mut Numbering<'a>,
) -> Result<Paragraph, String> {
    let level = list.depth.saturating_sub(1);
    let object_styles = env.object.for_level(env.sheet, level);
    let mut styles = paragraph_chain(env.sheet, env.scope, node);
    styles.extend(object_styles.iter().copied());
    let mut para_props = Props::new();
    para_props.extend_styles(&styles, PARAGRAPH);
    let mut text_props = Props::new();
    text_props.extend_styles(&styles, TEXT);
    env.payload.step(
        styles
            .len()
            .saturating_add(crate::odp_style::props_cost(&para_props))
            .saturating_add(crate::odp_style::props_cost(&text_props)),
    )?;

    let alignment = match para_props.get(NS_FO, "text-align") {
        Some("center") => "ctr",
        Some("end" | "right") => "r",
        Some("justify") => {
            if para_props.get(NS_FO, "text-align-last") == Some("center") {
                "ctr"
            } else {
                "just"
            }
        }
        _ => "l",
    };
    let rtl = matches!(
        para_props.get(NS_STYLE, "writing-mode"),
        Some("rl-tb" | "rl")
    );
    let alignment = if rtl && para_props.get(NS_FO, "text-align") == Some("start") {
        "r"
    } else {
        alignment
    };
    let length = |name: &str| para_props.get(NS_FO, name).and_then(length_emu);
    let margin_left = length("margin-left").unwrap_or(0.0);
    let margin_right = length("margin-right").unwrap_or(0.0);
    let text_indent = length("text-indent").unwrap_or(0.0);
    let space_before = length("margin-top").map(|v| (v / 127.0).round() as i64);
    let space_after = length("margin-bottom").map(|v| (v / 127.0).round() as i64);
    let space_line = para_props
        .get(NS_FO, "line-height")
        .and_then(|v| {
            if v == "normal" {
                None
            } else if let Some(p) = percent(v) {
                Some(SpaceLine::Pct { val: p * 100_000.0 })
            } else {
                length_emu(v).map(|e| SpaceLine::Pts {
                    val: e / EMU_PER_PT,
                })
            }
        })
        .or_else(|| {
            para_props
                .get(NS_STYLE, "line-height-at-least")
                .and_then(length_emu)
                .map(|e| SpaceLine::Pts {
                    val: e / EMU_PER_PT,
                })
        });

    // Lists: the label and the indents (§19.880 text:space-before,
    // §19.866 text:min-label-width), added to the paragraph's own margins
    // as LibreOffice's presentation engine does.
    let list_style_node = (list.depth > 0)
        .then(|| list_style(list, env, &styles))
        .flatten();
    let level_node = match list_style_node {
        Some(ls) => list_level(ls, list.depth, env.payload)?,
        None => None,
    };
    let mut mar_l = margin_left;
    let mut indent = text_indent;
    let mut bullet = Bullet::None;
    if let Some(level_node) = level_node {
        let props = level_node
            .children()
            .find(|c| xml::is(*c, NS_STYLE, "list-level-properties"));
        let get = |name: &str| {
            props
                .and_then(|p| p.attribute((NS_TEXT, name)))
                .and_then(length_emu)
        };
        let alignment_mode = props
            .and_then(|p| p.attribute((NS_TEXT, "list-level-position-and-space-mode")))
            == Some("label-alignment");
        if alignment_mode {
            let label = props.and_then(|p| {
                p.children()
                    .find(|c| xml::is(*c, NS_STYLE, "list-level-label-alignment"))
            });
            let la = |name: &str| {
                label
                    .and_then(|l| l.attribute((NS_FO, name)))
                    .and_then(length_emu)
            };
            if length("margin-left").is_none() {
                mar_l = la("margin-left").unwrap_or(0.0);
            }
            if length("text-indent").is_none() {
                indent = la("text-indent").unwrap_or(0.0);
            }
        } else {
            let space_before = get("space-before").unwrap_or(0.0);
            let min_label = get("min-label-width").unwrap_or(0.0);
            mar_l = space_before + min_label + margin_left;
            indent = -min_label + text_indent;
        }
        // Only an item's labelled paragraph takes part in the numbering; the
        // item's further paragraphs and list headers keep the sequence.
        if list.labelled {
            // Only a numbered level joins a consecutive sequence; a bullet
            // or picture item keeps its own level (and the sequence).
            let counted = if xml::is(level_node, NS_TEXT, "list-level-style-number") {
                consecutive_context(level_node, list)
            } else {
                list.clone()
            };
            let (label, number) = list_label(level_node, &counted, env, media, counters)?;
            bullet = label;
            counters.record(&counted, number);
        }
    } else if list.labelled {
        counters.record(list, None);
    }

    let tab_stops = para_props
        .child(NS_STYLE, "tab-stops")
        .map(|tabs| {
            tabs.children()
                .filter(|t| xml::is(*t, NS_STYLE, "tab-stop"))
                .filter_map(|t| {
                    let pos = t.attribute((NS_STYLE, "position")).and_then(length_emu)?;
                    let algn = match t.attribute((NS_STYLE, "type")) {
                        Some("center") => "ctr",
                        Some("right") => "r",
                        Some("char") => "dec",
                        _ => "l",
                    };
                    // Presentation tab positions are relative to the indent.
                    Some(TabStop {
                        pos: (pos + mar_l).round() as i64,
                        algn: algn.to_string(),
                    })
                })
                .collect()
        })
        .unwrap_or_default();
    let def_tab_sz = para_props
        .get(NS_STYLE, "tab-stop-distance")
        .or_else(|| para_props.get(NS_LOEXT, "tab-stop-distance"))
        .and_then(length_emu)
        .filter(|v| *v > 0.0)
        .map(|v| v.round() as i64);

    let base_run = run_properties(&text_props, env);
    // Its copies (the end run and the paragraph defaults), before they are made.
    env.payload
        .charge(record_bytes(&base_run)?.saturating_mul(2))?;
    let mut runs = Vec::new();
    let mut sink = SlideSink {
        stack: Vec::new(),
        links: vec![None],
        levels: vec![None],
        paragraph_styles: &styles,
        env,
        runs: &mut runs,
    };
    walk_paragraph(node, &mut sink, &mut Collapse::default(), 0)?;
    let end_run = base_run.clone();
    // A label's colour/size follow the first run when the list does not set them.
    let paragraph = Paragraph {
        alignment: alignment.to_string(),
        mar_l: mar_l.round() as i64,
        mar_r: margin_right.round() as i64,
        indent: indent.round() as i64,
        space_before,
        space_after,
        space_before_pct: None,
        space_after_pct: None,
        space_line,
        lvl: level.min(8) as u32,
        bullet,
        def_font_size: base_run.font_size,
        def_color: base_run.color.clone(),
        def_bold: base_run.bold,
        def_italic: base_run.italic,
        def_font_family: base_run.font_family.clone(),
        tab_stops,
        def_tab_sz,
        rtl,
        ea_ln_brk: true,
        font_algn: None,
        runs: Vec::new(),
        math_fallbacks: Vec::new(),
        end_run_properties: Some(Box::new(end_run)),
    };
    // The paragraph record itself (its runs were charged as they were read).
    env.payload.charge(record_bytes(&paragraph)?)?;
    Ok(Paragraph { runs, ..paragraph })
}

/// List numbering state of one text body (§19.785, §19.786): per list
/// instance and depth, the number of the latest labelled item at that depth
/// that no shallower item of the instance has ended (`Some(None)`: a
/// labelled item without a number, which restarts the sequence). Paragraphs
/// outside lists and of other instances do not affect an instance. The last
/// number of each level and of each identified list is also kept for
/// explicit continuations.
#[derive(Default)]
struct Numbering<'a> {
    instances: HashMap<usize, Vec<Option<Option<u32>>>>,
    /// Last number and list style per (instance, depth), not ended by
    /// shallower items.
    last_at_level: HashMap<(usize, usize), (u32, Option<Nd<'a>>)>,
    /// Last number per list `xml:id`.
    last_of_list: HashMap<&'a str, u32>,
}

impl<'a> Numbering<'a> {
    /// The number the next item of `list` continues from, if any.
    fn previous(&self, list: &ListContext<'a>) -> Option<u32> {
        let instance = list.instance?;
        if list.first_in_list {
            match list.resume {
                Resume::List(id) => return self.last_of_list.get(id).copied(),
                Resume::Preceding => {
                    return self
                        .last_at_level
                        .get(&(instance, list.depth))
                        .filter(|(_, style)| *style == list.numbering_style)
                        .map(|(n, _)| *n)
                }
                Resume::Restart => return None,
                Resume::After(n) => return Some(n),
                Resume::None => {}
            }
        }
        self.instances
            .get(&instance)
            .and_then(|levels| levels.get(list.depth).copied().flatten())
            .flatten()
    }

    fn record(&mut self, list: &ListContext<'a>, number: Option<u32>) {
        let Some(instance) = list.instance else {
            return;
        };
        let levels = self.instances.entry(instance).or_default();
        if levels.len() <= list.depth {
            levels.resize(list.depth + 1, None);
        }
        levels[list.depth] = Some(number);
        // A shallower or equal item ends the deeper sequences.
        levels.truncate(list.depth + 1);
        if let Some(n) = number {
            self.last_at_level
                .insert((instance, list.depth), (n, list.numbering_style));
            if let Some(id) = list.list_id {
                self.last_of_list.insert(id, n);
            }
        }
    }
}

/// The numbering style a `text:list` declares, else the one the paragraph
/// styles of its first paragraph select (searched through leading nested
/// lists; every node inspected is charged).
fn effective_list_style<'a>(
    instances: &ListInstances<'a, '_>,
    list: Nd<'a>,
    outer: &ListContext<'a>,
) -> Result<Option<Nd<'a>>, String> {
    let (sheet, scope) = (instances.sheet, instances.scope);
    // The nearest list naming a style, as rendering resolves it.
    let named = if list.attribute((NS_TEXT, "style-name")).is_some() {
        Some(list)
    } else {
        outer.style
    };
    let level = outer.depth;
    // Each resolution scans the cascade (style children, their graphic
    // properties' children): charged, found or not.
    let cascade = |first: Option<Nd<'a>>| -> Result<Option<Nd<'a>>, String> {
        let name_cost = named
            .and_then(|l| l.attribute((NS_TEXT, "style-name")))
            .map_or(0, |n| n.len() / 64);
        let mut styles = first
            .map(|p| paragraph_chain(sheet, scope, p))
            .unwrap_or_default();
        styles.extend(instances.object.for_level(sheet, level));
        instances.payload.step(
            1usize.saturating_add(name_cost).saturating_add(
                styles
                    .len()
                    .saturating_mul(2 * crate::odp_style::STYLE_SCAN_COST),
            ),
        )?;
        Ok(resolve_list_style(sheet, scope, named, &styles))
    };
    if named.is_some_and(|l| l.attribute((NS_TEXT, "style-name")).is_some()) {
        return cascade(None);
    }
    let first = |parent: Nd<'a>, pick: &dyn Fn(Nd<'a>) -> bool| -> Result<Option<Nd<'a>>, String> {
        for child in parent.children() {
            instances.payload.step(1)?;
            if child.is_element() && pick(child) {
                return Ok(Some(child));
            }
        }
        Ok(None)
    };
    let mut node = list;
    for _ in 0..MAX_SPAN_DEPTH {
        let Some(item) = first(node, &|c| {
            xml::is(c, NS_TEXT, "list-item") || xml::is(c, NS_TEXT, "list-header")
        })?
        else {
            return Ok(None);
        };
        let Some(block) = first(item, &|c| {
            xml::is(c, NS_TEXT, "p") || xml::is(c, NS_TEXT, "h") || xml::is(c, NS_TEXT, "list")
        })?
        else {
            return Ok(None);
        };
        if xml::is(block, NS_TEXT, "list") {
            node = block;
            continue;
        }
        return cascade(Some(block));
    }
    cascade(None)
}

/// The styles of the object holding a text (the frame of a text box, else
/// the element itself), for continuation resolution outside rendering:
/// its presentation or graphic style chain, its text style and the default
/// graphic style (a placeholder's per-level outline styles are not known
/// here).
fn container_object_styles<'a>(
    sheet: &'a Sheet<'a>,
    scope: Scope,
    container: Nd<'a>,
    payload: &TextPayload<'_>,
) -> Result<ObjectStyles<'a>, String> {
    // A table cell's text: the cell's style cascade (its own or its row's
    // default cell style, then the table-cell default) before the styles
    // of the frame holding the table, as the table reader builds them
    // (a table template's region styles are not known here).
    if xml::is(container, crate::xml::NS_TABLE, "table-cell") {
        let frame = container
            .ancestors()
            .find(|a| xml::is(*a, NS_DRAW, "frame"))
            .unwrap_or(container);
        let mut styles = if frame == container {
            ObjectStyles::default()
        } else {
            container_object_styles(sheet, scope, frame, payload)?
        };
        // §16.2: the cell's style, then its row's and its column's default
        // cell styles, as the table reader cascades them.
        let mut chain = Vec::new();
        for name in [
            container.attribute((crate::xml::NS_TABLE, "style-name")),
            container
                .parent_element()
                .and_then(|row| row.attribute((crate::xml::NS_TABLE, "default-cell-style-name"))),
            crate::odp_table::column_default_style(container, payload)?,
        ]
        .into_iter()
        .flatten()
        {
            chain.extend(sheet.chain(scope, "table-cell", name)?);
        }
        chain.extend(crate::odp_table::cell_template_style(
            sheet, scope, container, payload,
        )?);
        chain.extend(sheet.default_style("table-cell"));
        chain.extend(styles.chain.iter().copied());
        styles.chain = chain;
        styles.outline_master = None;
        styles.automatic = None;
        return Ok(styles);
    }
    let object = if xml::is(container, NS_DRAW, "text-box") {
        container.parent_element().unwrap_or(container)
    } else {
        container
    };
    let (family, name) = match object.attribute((NS_PRESENTATION, "style-name")) {
        Some(name) => ("presentation", Some(name)),
        None => ("graphic", object.attribute((NS_DRAW, "style-name"))),
    };
    let chain = match name {
        Some(name) => sheet.chain(scope, family, name)?,
        None => Vec::new(),
    };
    let text_style = object
        .attribute((NS_DRAW, "text-style-name"))
        .map(|n| sheet.chain(scope, "paragraph", n))
        .transpose()?
        .unwrap_or_default();
    Ok(ObjectStyles {
        chain,
        automatic: None,
        outline_master: None,
        text_style,
        defaults: sheet.default_style("graphic").into_iter().collect(),
    })
}

/// The label of a list item (§16.31–§16.33 list level styles).
fn list_label<'a>(
    level_node: Nd<'a>,
    list: &ListContext<'a>,
    env: &TextEnv<'a, '_>,
    media: &mut Media<'_>,
    counters: &Numbering<'a>,
) -> Result<(Bullet, Option<u32>), String> {
    let level_text = level_node
        .children()
        .find(|c| xml::is(*c, NS_STYLE, "text-properties"));
    let label_color = level_text.and_then(|t| {
        if t.attribute((NS_STYLE, "use-window-font-color")) == Some("true") {
            None
        } else {
            t.attribute((NS_FO, "color")).and_then(color)
        }
    });
    let label_size = level_text
        .and_then(|t| t.attribute((NS_FO, "font-size")))
        .and_then(|v| {
            percent(v)
                .map(|p| (Some(p * 100.0), None))
                .or_else(|| length_emu(v).map(|e| (None, Some(e / EMU_PER_PT))))
        })
        .unwrap_or((None, None));
    let label_font = level_text.and_then(|t| {
        t.attribute((NS_STYLE, "font-name"))
            .and_then(|f| env.sheet.font_family(f).map(str::to_string))
            .or_else(|| {
                t.attribute((NS_FO, "font-family"))
                    .map(crate::odf_props::unquote_family)
            })
    });
    if xml::is(level_node, NS_TEXT, "list-level-style-bullet") {
        let ch = level_node.attribute((NS_TEXT, "bullet-char")).unwrap_or("");
        if ch.is_empty() {
            return Ok((Bullet::None, None));
        }
        let size_pct = level_node
            .attribute((NS_TEXT, "bullet-relative-size"))
            .and_then(percent)
            .map(|p| p * 100.0)
            .or(label_size.0);
        return Ok((
            Bullet::Char {
                ch: ch.to_string(),
                color: label_color,
                size_pct,
                size_pts: label_size.1,
                font_family: label_font,
            },
            None,
        ));
    }
    if xml::is(level_node, NS_TEXT, "list-level-style-image") {
        let Some(href) = level_node.attribute((NS_XLINK, "href")) else {
            return Ok((Bullet::None, None));
        };
        let Some((path, resource)) = media.admit(env.base, href) else {
            return Ok((Bullet::None, None));
        };
        let props = level_node
            .children()
            .find(|c| xml::is(*c, NS_STYLE, "list-level-properties"));
        let height = props
            .and_then(|p| p.attribute((NS_FO, "height")))
            .and_then(length_emu);
        return Ok((
            Bullet::Blip {
                image_path: path,
                mime_type: resource.mime.to_string(),
                size_pct: None,
                size_pts: height.map(|h| h / EMU_PER_PT),
            },
            None,
        ));
    }
    if xml::is(level_node, NS_TEXT, "list-level-style-number") {
        // The earlier items of the same list instance (§19.785, §19.786):
        // an equal-depth numbered item continues the sequence; a shallower
        // item ends it (a nested list restarts under each parent item).
        let start = level_node
            .attribute((NS_TEXT, "start-value"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1);
        let number = match (list.start_value, counters.previous(list)) {
            (Some(value), _) => value,
            (None, Some(n)) => n.saturating_add(1),
            (None, None) => start,
        };
        let format = level_node
            .attribute((NS_STYLE, "num-format"))
            .unwrap_or("1");
        let prefix = level_node.attribute((NS_STYLE, "num-prefix")).unwrap_or("");
        let suffix = level_node.attribute((NS_STYLE, "num-suffix")).unwrap_or("");
        let label = format!("{prefix}{}{suffix}", format_number(number, format));
        if label.is_empty() {
            return Ok((Bullet::None, Some(number)));
        }
        return Ok((
            Bullet::Char {
                ch: label,
                color: label_color,
                size_pct: label_size.0,
                size_pts: label_size.1,
                font_family: label_font,
            },
            Some(number),
        ));
    }
    Ok((Bullet::None, None))
}

/// §19.340 style:num-format.
pub fn format_number(n: u32, format: &str) -> String {
    match format {
        "" => String::new(),
        "a" | "A" => {
            let mut v = n.max(1);
            let mut letters = Vec::new();
            while v > 0 {
                v -= 1;
                letters.push((b'a' + (v % 26) as u8) as char);
                v /= 26;
            }
            let s: String = letters.into_iter().rev().collect();
            if format == "A" {
                s.to_ascii_uppercase()
            } else {
                s
            }
        }
        "i" | "I" => {
            let s = roman(n);
            if format == "i" {
                s.to_ascii_lowercase()
            } else {
                s
            }
        }
        _ => n.to_string(),
    }
}

fn roman(mut n: u32) -> String {
    if n == 0 || n >= 4000 {
        return n.to_string();
    }
    let table = [
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
    for (value, text) in table {
        while n >= value {
            out.push_str(text);
            n -= value;
        }
    }
    out
}

/// Script class of a character for the western / Asian / complex property
/// triples (§20.x `*-asian`, `*-complex`). Weak characters return `None`.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Script {
    Western,
    Asian,
    Complex,
}

fn script_of(c: char) -> Option<Script> {
    let u = c as u32;
    if c.is_whitespace() || c.is_ascii_punctuation() || c.is_ascii_digit() {
        return None;
    }
    if (0x0590..=0x08FF).contains(&u)
        || (0x0900..=0x0DFF).contains(&u)
        || (0x0E00..=0x0EFF).contains(&u)
        || (0xFB1D..=0xFDFF).contains(&u)
        || (0xFE70..=0xFEFF).contains(&u)
    {
        return Some(Script::Complex);
    }
    if (0x1100..=0x11FF).contains(&u)
        || (0x2E80..=0x9FFF).contains(&u)
        || (0xA960..=0xA97F).contains(&u)
        || (0xAC00..=0xD7FF).contains(&u)
        || (0xF900..=0xFAFF).contains(&u)
        || (0xFE30..=0xFE4F).contains(&u)
        || (0xFF00..=0xFFEF).contains(&u)
        || (0x20000..=0x3FFFF).contains(&u)
    {
        return Some(Script::Asian);
    }
    Some(Script::Western)
}

/// Run formatting resolved from a text-properties cascade (western values;
/// Asian and complex script values are applied per script portion).
#[derive(Clone)]
struct RunStyle {
    data: TextRunData,
    /// `fo:text-transform` lowercase / capitalize, applied to the text.
    transform: Option<&'static str>,
    size: [Option<f64>; 3],
    bold: [Option<bool>; 3],
    italic: [Option<bool>; 3],
}

fn font_size(props: &Props<'_>, attr: (&str, &str)) -> Option<f64> {
    // Percentages scale the inherited size: resolve from the farthest.
    let mut size: Option<f64> = None;
    for node in props.nodes.iter().rev() {
        if let Some(v) = node.attribute(attr) {
            if let Some(p) = percent(v) {
                size = size.map(|s| s * p);
            } else if let Some(e) = length_emu(v) {
                size = Some(e / EMU_PER_PT);
            }
        }
    }
    size
}

fn weight(value: Option<&str>) -> Option<bool> {
    let v = value?;
    Some(match v {
        "bold" => true,
        "normal" => false,
        other => other.parse::<u32>().map(|w| w >= 600).unwrap_or(false),
    })
}

fn posture(value: Option<&str>) -> Option<bool> {
    value.map(|v| v == "italic" || v == "oblique")
}

fn face(
    props: &Props<'_>,
    sheet: &Sheet<'_>,
    name_attr: &str,
    family_attr: (&str, &str),
) -> Option<String> {
    for node in &props.nodes {
        if let Some(face) = node.attribute((NS_STYLE, name_attr)) {
            return Some(
                sheet
                    .font_family(face)
                    .map(str::to_string)
                    .unwrap_or_else(|| face.to_string()),
            );
        }
        if let Some(family) = node.attribute(family_attr) {
            return Some(crate::odf_props::unquote_family(family));
        }
    }
    None
}

fn run_style(props: &Props<'_>, env: &TextEnv<'_, '_>) -> RunStyle {
    let sheet = env.sheet;
    let size = [
        font_size(props, (NS_FO, "font-size")),
        font_size(props, (NS_STYLE, "font-size-asian")),
        font_size(props, (NS_STYLE, "font-size-complex")),
    ];
    let bold = [
        weight(props.get(NS_FO, "font-weight")),
        weight(props.get(NS_STYLE, "font-weight-asian")),
        weight(props.get(NS_STYLE, "font-weight-complex")),
    ];
    let italic = [
        posture(props.get(NS_FO, "font-style")),
        posture(props.get(NS_STYLE, "font-style-asian")),
        posture(props.get(NS_STYLE, "font-style-complex")),
    ];
    // Colour: the nearest of fo:color and style:use-window-font-color wins;
    // an automatic colour is black, or white on a dark background.
    let mut colour: Option<String> = None;
    for node in &props.nodes {
        let auto = node.attribute((NS_STYLE, "use-window-font-color")) == Some("true");
        if auto {
            colour = Some(match &env.background {
                Some(bg) if is_dark(bg) => "FFFFFF".into(),
                _ => "000000".into(),
            });
            break;
        }
        if let Some(c) = node.attribute((NS_FO, "color")).and_then(color) {
            let opacity = node
                .attribute((NS_LOEXT, "opacity"))
                .and_then(percent)
                .unwrap_or(1.0);
            colour = Some(with_alpha(&c, opacity));
            break;
        }
    }
    let underline_style = props
        .get(NS_STYLE, "text-underline-style")
        .unwrap_or("none");
    let underline_type = props
        .get(NS_STYLE, "text-underline-type")
        .unwrap_or("single");
    let underline_bold = props.get(NS_STYLE, "text-underline-width") == Some("bold");
    let underline = underline_style != "none" && underline_type != "none";
    let underline_name = if !underline {
        None
    } else {
        let base = match (underline_style, underline_type) {
            (_, "double") => "dbl",
            ("dotted", _) => {
                if underline_bold {
                    "dottedHeavy"
                } else {
                    "dotted"
                }
            }
            ("dash", _) => {
                if underline_bold {
                    "dashHeavy"
                } else {
                    "dash"
                }
            }
            ("long-dash", _) => {
                if underline_bold {
                    "dashLongHeavy"
                } else {
                    "dashLong"
                }
            }
            ("dot-dash", _) => {
                if underline_bold {
                    "dotDashHeavy"
                } else {
                    "dotDash"
                }
            }
            ("dot-dot-dash", _) => {
                if underline_bold {
                    "dotDotDashHeavy"
                } else {
                    "dotDotDash"
                }
            }
            ("wave", _) => {
                if underline_type == "double" {
                    "wavyDbl"
                } else if underline_bold {
                    "wavyHeavy"
                } else {
                    "wavy"
                }
            }
            _ => {
                if underline_bold {
                    "heavy"
                } else {
                    "sng"
                }
            }
        };
        (base != "sng").then(|| base.to_string())
    };
    let underline_color = props
        .get(NS_STYLE, "text-underline-color")
        .filter(|v| *v != "font-color")
        .and_then(color);
    let strike_style = props
        .get(NS_STYLE, "text-line-through-style")
        .unwrap_or("none");
    let strike_type = props
        .get(NS_STYLE, "text-line-through-type")
        .unwrap_or("single");
    let strikethrough = strike_style != "none" && strike_type != "none";
    let caps = match (
        props.get(NS_FO, "text-transform"),
        props.get(NS_FO, "font-variant"),
    ) {
        (Some("uppercase"), _) => Some("all".to_string()),
        (_, Some("small-caps")) => Some("small".to_string()),
        _ => None,
    };
    let letter_spacing = props
        .get(NS_FO, "letter-spacing")
        .filter(|v| *v != "normal")
        .and_then(length_emu)
        .map(|e| e / EMU_PER_PT)
        .filter(|v| *v != 0.0);
    let shadow = props
        .get(NS_FO, "text-shadow")
        .filter(|v| *v != "none")
        .map(|v| {
            let parts: Vec<f64> = v
                .split_whitespace()
                .filter_map(length_emu)
                .take(2)
                .collect();
            let (ox, oy) = (
                parts.first().copied().unwrap_or(1.0 * EMU_PER_PT),
                parts.get(1).copied().unwrap_or(1.0 * EMU_PER_PT),
            );
            Shadow {
                color: v
                    .split_whitespace()
                    .find_map(color)
                    .unwrap_or_else(|| "808080".into()),
                alpha: 1.0,
                blur: 0,
                dist: ox.hypot(oy).round() as i64,
                dir: oy.atan2(ox).to_degrees().rem_euclid(360.0),
                sx: None,
                sy: None,
                kx: None,
                ky: None,
                algn: None,
                rot_with_shape: None,
            }
        });
    let outline_on = props.get(NS_STYLE, "text-outline") == Some("true");
    let highlight = props
        .get(NS_FO, "background-color")
        .filter(|v| *v != "transparent")
        .and_then(color);
    let lang = props
        .get(NS_FO, "language")
        .filter(|l| *l != "none")
        .map(|l| match props.get(NS_FO, "country") {
            Some(c) if c != "none" => format!("{l}-{c}"),
            _ => l.to_string(),
        });
    let (baseline, size_factor) = match props.get(NS_STYLE, "text-position") {
        Some(v) => text_position(v),
        None => (None, 1.0),
    };
    let data = TextRunData {
        text: String::new(),
        bold: bold[0],
        italic: italic[0],
        underline,
        underline_style: underline_name,
        underline_color,
        underline_fill: None,
        underline_line: None,
        underline_line_no_fill: false,
        strikethrough,
        strike_double: strikethrough && strike_type == "double",
        font_size: size[0].map(|s| s * size_factor),
        color: colour.clone(),
        pattern_fill: None,
        glyph_fill: None,
        no_fill: outline_on,
        font_family: face(props, sheet, "font-name", (NS_FO, "font-family")),
        font_family_ea: face(
            props,
            sheet,
            "font-name-asian",
            (NS_STYLE, "font-family-asian"),
        ),
        font_family_cs: face(
            props,
            sheet,
            "font-name-complex",
            (NS_STYLE, "font-family-complex"),
        ),
        font_family_sym: None,
        lang,
        alt_lang: None,
        baseline,
        caps,
        letter_spacing,
        field_type: None,
        hyperlink: None,
        hyperlink_uses_text_fill: false,
        hyperlink_action: None,
        hyperlink_mouse_over: None,
        hyperlink_mouse_over_action: None,
        shadow,
        reflection: None,
        outline: outline_on.then_some(TextOutline {
            width: 9525,
            color: colour,
            fill: None,
        }),
        highlight,
        character_attributes: BTreeMap::new(),
        character_child_attributes: BTreeMap::new(),
    };
    let transform = match props.get(NS_FO, "text-transform") {
        Some("lowercase") => Some("lowercase"),
        Some("capitalize") => Some("capitalize"),
        _ => None,
    };
    RunStyle {
        data,
        transform,
        size: size.map(|s| s.map(|v| v * size_factor)),
        bold,
        italic,
    }
}

/// `style:text-position` (§20.374): "super"/"sub"/percent raise, then the
/// relative font height (default 58% for super/sub). The PPTX renderer
/// draws a raised run at 65% of its size, so the size is pre-scaled to keep
/// the authored relative height.
fn text_position(value: &str) -> (Option<i32>, f64) {
    let mut parts = value.split_whitespace();
    let raise = match parts.next() {
        Some("super") => 0.33,
        Some("sub") => -0.33,
        Some(p) => percent(p).unwrap_or(0.0),
        None => 0.0,
    };
    let relative = parts.next().and_then(percent).unwrap_or(
        if value.starts_with("super") || value.starts_with("sub") {
            0.58
        } else {
            1.0
        },
    );
    if raise == 0.0 {
        return (None, relative);
    }
    let factor = relative / 0.65;
    ((raise / factor * 100_000.0).round() as i32).into_some_nonzero(factor)
}

trait IntoSomeNonzero {
    fn into_some_nonzero(self, factor: f64) -> (Option<i32>, f64);
}
impl IntoSomeNonzero for i32 {
    fn into_some_nonzero(self, factor: f64) -> (Option<i32>, f64) {
        if self == 0 {
            (None, 1.0)
        } else {
            (Some(self), factor)
        }
    }
}

fn run_properties(props: &Props<'_>, env: &TextEnv<'_, '_>) -> TextRunData {
    run_style(props, env).data
}

/// How an element inside a paragraph is read (ODF 1.3 Part 3 §6, §7).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Inline {
    /// A text container (`text:span`, `text:a`, `text:meta`,
    /// `text:meta-field`, an unknown element, …): its character data
    /// collapses with the paragraph's (§6.1.2).
    Container,
    /// `text:ruby` (§6.4): its `text:ruby-base` is a container; the ruby
    /// text is not shown.
    Ruby,
    /// A §7 field: its stored representation is shown as written.
    Field,
    /// `text:s` (§6.1.3): this many significant spaces.
    Space(usize),
    /// `text:tab` (§6.1.4).
    Tab,
    /// `text:line-break` (§6.1.5).
    LineBreak,
    /// Not shown: notes, marks, change tracking, annotations, scripts,
    /// drawings and object titles.
    Ignored,
}

/// §7 field elements (text namespace), whose stored text is kept verbatim.
const TEXT_FIELDS: &[&str] = &[
    // §7.3 document fields
    "date",
    "time",
    "page-number",
    "page-continuation",
    "sender-firstname",
    "sender-lastname",
    "sender-initials",
    "sender-title",
    "sender-position",
    "sender-email",
    "sender-phone-private",
    "sender-fax",
    "sender-company",
    "sender-phone-work",
    "sender-street",
    "sender-city",
    "sender-postal-code",
    "sender-country",
    "sender-state-or-province",
    "author-name",
    "author-initials",
    "chapter",
    "file-name",
    "template-name",
    "sheet-name",
    // §7.4 variable fields
    "variable-set",
    "variable-get",
    "variable-input",
    "user-field-get",
    "user-field-input",
    "sequence",
    "expression",
    "text-input",
    "drop-down",
    // §7.5 metadata fields
    "initial-creator",
    "creation-date",
    "creation-time",
    "description",
    "user-defined",
    "print-time",
    "print-date",
    "printed-by",
    "title",
    "subject",
    "keywords",
    "editing-cycles",
    "editing-duration",
    "modification-time",
    "modification-date",
    "creator",
    "page-count",
    "paragraph-count",
    "word-count",
    "character-count",
    "table-count",
    "image-count",
    "object-count",
    // §7.6 database fields
    "database-display",
    "database-next",
    "database-row-select",
    "database-row-number",
    "database-name",
    // §7.7 other fields
    "page-variable-set",
    "page-variable-get",
    "placeholder",
    "conditional-text",
    "hidden-text",
    "reference-ref",
    "sequence-ref",
    "bookmark-ref",
    "note-ref",
    "execute-macro",
    "hidden-paragraph",
    "dde-connection",
    "measure",
    "table-formula",
];

/// Inline elements that are not shown (text namespace).
const TEXT_IGNORED: &[&str] = &[
    "note",
    "bookmark",
    "bookmark-start",
    "bookmark-end",
    "soft-page-break",
    "reference-mark",
    "reference-mark-start",
    "reference-mark-end",
    "change-start",
    "change-end",
    "change",
    "toc-mark",
    "toc-mark-start",
    "toc-mark-end",
    "alphabetical-index-mark",
    "alphabetical-index-mark-start",
    "alphabetical-index-mark-end",
    "user-index-mark",
    "user-index-mark-start",
    "user-index-mark-end",
    // Script source, not displayed text (§7.7.9).
    "script",
    // The choices of a `text:drop-down` (§7.4.17); its text is the value.
    "label",
];

/// The 3D-scene namespace (§10.5): a drawing, as `draw:*` shapes are.
const NS_DR3D: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";

/// The class of an inline element (§6, §7).
pub(crate) fn classify(node: Nd<'_>) -> Inline {
    let name = node.tag_name().name();
    match node.tag_name().namespace() {
        Some(NS_TEXT) => match name {
            "span" | "a" | "meta" | "meta-field" | "ruby-base" => Inline::Container,
            "ruby" => Inline::Ruby,
            "s" => Inline::Space(
                node.attribute((NS_TEXT, "c"))
                    .and_then(|c| c.parse::<usize>().ok())
                    .unwrap_or(1)
                    .clamp(1, 1024),
            ),
            "tab" => Inline::Tab,
            "line-break" => Inline::LineBreak,
            _ if TEXT_FIELDS.contains(&name) => Inline::Field,
            _ if TEXT_IGNORED.contains(&name) => Inline::Ignored,
            _ => Inline::Container,
        },
        Some(NS_PRESENTATION) if matches!(name, "footer" | "header" | "date-time") => Inline::Field,
        Some(crate::xml::NS_OFFICE) if matches!(name, "annotation" | "annotation-end") => {
            Inline::Ignored
        }
        // Drawings anchored in the text (frames, shapes, 3D scenes) and the
        // titles and descriptions of objects are not paragraph text.
        Some(NS_DRAW) | Some(NS_DR3D) => Inline::Ignored,
        Some(crate::xml::NS_SVG) if matches!(name, "title" | "desc") => Inline::Ignored,
        _ => Inline::Container,
    }
}

/// Receives the content of one paragraph, read by [`walk_paragraph`]: the
/// styled runs of slide text, or the plain text of speaker notes.
trait InlineSink<'a> {
    /// A held space: what it needs to be emitted later.
    type Space;
    /// Charge one visited node.
    fn visit(&mut self) -> Result<(), String>;
    /// Charge `units` of work.
    fn step(&mut self, units: usize) -> Result<(), String>;
    /// A container opens / closes.
    fn enter(&mut self, element: Nd<'a>);
    fn leave(&mut self);
    /// Text of the current container (already collapsed, or verbatim).
    fn text(&mut self, piece: &str) -> Result<(), String>;
    /// Hold a collapsed space here; `after_text`: right after this
    /// container's own text.
    fn hold_space(&mut self, after_text: bool) -> Result<Self::Space, String>;
    fn put_space(&mut self, space: Self::Space) -> Result<(), String>;
    /// `text:s`, `text:tab` or `text:line-break`.
    fn preserved(&mut self, class: Inline) -> Result<(), String>;
    /// A field the sink generates itself; `false` shows its stored text.
    fn generated(&mut self, field: Nd<'a>) -> Result<bool, String>;
}

/// White-space collapsing state of one paragraph (§6.1.2): whether content
/// has started, a space waiting for following content, and whether the
/// text is inside a field (kept verbatim).
struct Collapse<S> {
    started: bool,
    pending: Option<S>,
    verbatim: bool,
}

impl<S> Default for Collapse<S> {
    fn default() -> Self {
        Collapse {
            started: false,
            pending: None,
            verbatim: false,
        }
    }
}

/// Read the inline content of a paragraph for both slide text and notes:
/// §6.1.2 white space collapses over the whole paragraph (a space carried
/// from an earlier node is emitted only when content follows; leading and
/// trailing spaces are dropped), preserved white space and fields are
/// content, and field text is kept as stored.
fn walk_paragraph<'a, K: InlineSink<'a>>(
    node: Nd<'a>,
    sink: &mut K,
    ws: &mut Collapse<K::Space>,
    depth: usize,
) -> Result<(), String> {
    if depth > MAX_SPAN_DEPTH {
        return Err(crate::error::invalid("text spans are nested too deeply"));
    }
    for child in node.children() {
        sink.visit()?;
        if child.is_text() {
            let text = child.text().unwrap_or("");
            sink.step(text.len() / 64)?;
            if ws.verbatim {
                sink.text(text)?;
                continue;
            }
            let mut carried = ws.pending.take();
            let mut pending_here = false;
            let mut piece = String::new();
            for c in text.chars() {
                if matches!(c, ' ' | '\t' | '\n' | '\r') {
                    if ws.started {
                        pending_here = true;
                    }
                    continue;
                }
                if let Some(space) = carried.take() {
                    sink.put_space(space)?;
                    pending_here = false;
                } else if pending_here {
                    piece.push(' ');
                    pending_here = false;
                }
                piece.push(c);
                ws.started = true;
            }
            let after_text = !piece.is_empty();
            sink.text(&piece)?;
            ws.pending = match carried {
                Some(space) => Some(space),
                None if pending_here => Some(sink.hold_space(after_text)?),
                None => None,
            };
            continue;
        }
        if !child.is_element() {
            continue;
        }
        match classify(child) {
            Inline::Container => {
                sink.enter(child);
                walk_paragraph(child, sink, ws, depth + 1)?;
                sink.leave();
            }
            Inline::Ruby => {
                if let Some(base) = child.children().find(|c| xml::is(*c, NS_TEXT, "ruby-base")) {
                    walk_paragraph(base, sink, ws, depth + 1)?;
                }
            }
            Inline::Ignored => {}
            class => {
                // Preserved white space and fields are content: a carried
                // space before them is kept.
                if let Some(space) = ws.pending.take() {
                    sink.put_space(space)?;
                }
                ws.started = true;
                if class != Inline::Field {
                    sink.preserved(class)?;
                } else if !sink.generated(child)? {
                    let outer = ws.verbatim;
                    ws.verbatim = true;
                    walk_paragraph(child, sink, ws, depth + 1)?;
                    ws.verbatim = outer;
                }
            }
        }
    }
    Ok(())
}

/// Slide text: styled runs, each container level resolving its run style
/// once (on first use) through the open containers and paragraph styles.
struct SlideSink<'a, 'e, 'x> {
    stack: Vec<Nd<'a>>,
    links: Vec<Option<&'a str>>,
    levels: Vec<Option<RunStyle>>,
    paragraph_styles: &'x [Nd<'a>],
    env: &'x TextEnv<'a, 'e>,
    runs: &'x mut Vec<TextRun>,
}

/// The run style of the innermost level, resolved once; resolving charges
/// the style records it reads (each chain node and the attributes of its
/// text properties).
fn level_style<'s, 'a>(
    levels: &'s mut [Option<RunStyle>],
    stack: &[Nd<'a>],
    paragraph_styles: &[Nd<'a>],
    env: &TextEnv<'a, '_>,
) -> Result<&'s RunStyle, String> {
    let level = levels.last_mut().expect("the paragraph level");
    if level.is_none() {
        let mut nodes: Vec<Nd<'a>> = Vec::new();
        for span in stack.iter().rev() {
            if let Some(name) = span.attribute((NS_TEXT, "style-name")) {
                // The reference is hashed for the lookup, found or not.
                env.payload.step(1 + name.len() / 64)?;
                if let Ok(chain) = env.sheet.chain(env.scope, "text", name) {
                    nodes.extend(chain);
                }
            }
        }
        nodes.extend(paragraph_styles.iter().copied());
        let mut props = Props::new();
        props.extend_styles(&nodes, TEXT);
        env.payload.step(
            nodes
                .len()
                .saturating_add(crate::odp_style::props_cost(&props)),
        )?;
        *level = Some(run_style(&props, env));
    }
    Ok(level.as_ref().expect("resolved above"))
}

impl<'a> SlideSink<'a, '_, '_> {
    fn link(&self) -> Option<&'a str> {
        self.links.last().copied().flatten()
    }

    fn style(&mut self) -> Result<&RunStyle, String> {
        level_style(
            &mut self.levels,
            &self.stack,
            self.paragraph_styles,
            self.env,
        )
    }

    fn push(&mut self, piece: &str) -> Result<(), String> {
        let link = self.link();
        let style = level_style(
            &mut self.levels,
            &self.stack,
            self.paragraph_styles,
            self.env,
        )?;
        push_text(self.runs, style, piece, link, self.env)
    }
}

impl<'a> InlineSink<'a> for SlideSink<'a, '_, '_> {
    /// The space's style and link, and whether it directly follows the last
    /// run (of its style).
    type Space = (RunStyle, Option<String>, bool);

    fn visit(&mut self) -> Result<(), String> {
        // Each node costs its style resolution through the open spans.
        self.env.payload.step(1 + self.stack.len())
    }

    fn step(&mut self, units: usize) -> Result<(), String> {
        self.env.payload.step(units)
    }

    fn enter(&mut self, element: Nd<'a>) {
        let link = if xml::is(element, NS_TEXT, "a") {
            element.attribute((NS_XLINK, "href")).or(self.link())
        } else {
            self.link()
        };
        self.stack.push(element);
        self.links.push(link);
        self.levels.push(None);
    }

    fn leave(&mut self) {
        self.stack.pop();
        self.links.pop();
        self.levels.pop();
    }

    fn text(&mut self, piece: &str) -> Result<(), String> {
        self.push(piece)
    }

    fn hold_space(&mut self, after_text: bool) -> Result<Self::Space, String> {
        let link = self.link().map(str::to_string);
        Ok((self.style()?.clone(), link, after_text))
    }

    /// Appended to the run it follows when that run has its style, else a
    /// run of its own.
    fn put_space(&mut self, (style, link, append): Self::Space) -> Result<(), String> {
        if append {
            if let Some(TextRun::Text(data)) = self.runs.last_mut() {
                self.env.payload.charge(1)?;
                data.text.push(' ');
                return Ok(());
            }
        }
        push_text(self.runs, &style, " ", link.as_deref(), self.env)
    }

    fn preserved(&mut self, class: Inline) -> Result<(), String> {
        match class {
            Inline::Space(count) => self.push(&" ".repeat(count)),
            Inline::Tab => self.push("\t"),
            _ => {
                let s = self.style()?;
                let run = TextRun::Break {
                    font_size: s.data.font_size,
                    font_family: s.data.font_family.clone(),
                    bold: s.data.bold,
                    italic: s.data.italic,
                    character_attributes: BTreeMap::new(),
                    character_child_attributes: BTreeMap::new(),
                };
                self.env.payload.charge(record_bytes(&run)?)?;
                self.runs.push(run);
                Ok(())
            }
        }
    }

    fn generated(&mut self, field: Nd<'a>) -> Result<bool, String> {
        let env = self.env;
        let name = field.tag_name().name();
        match (field.tag_name().namespace(), name) {
            (Some(NS_TEXT), "page-number") => {
                let format = field.attribute((NS_STYLE, "num-format")).unwrap_or("1");
                let mut s = self.style()?.clone();
                s.data.field_type = Some("slidenum".into());
                let text = format_number(env.slide_number as u32, format);
                push_text(self.runs, &s, &text, self.link(), env)?;
            }
            (Some(NS_TEXT), "page-count") => self.push(&env.slide_count.to_string())?,
            (Some(NS_PRESENTATION), _) => {
                let value = match name {
                    "footer" => env.footer.clone(),
                    "header" => env.header.clone(),
                    _ => env.date_time.clone(),
                };
                if let Some(text) = value {
                    self.push(&text)?;
                }
            }
            _ => return Ok(false),
        }
        Ok(true)
    }
}

/// Speaker notes: plain text within [`MAX_NOTES_BYTES`] (checked before
/// each piece is added).
struct NotesSink<'t> {
    text: &'t mut String,
}

impl NotesSink<'_> {
    fn push(&mut self, piece: &str) -> Result<(), String> {
        if self.text.len().saturating_add(piece.len()) > MAX_NOTES_BYTES {
            return Err("OpenDocument speaker notes exceed the text limit".to_string());
        }
        self.text.push_str(piece);
        Ok(())
    }
}

impl<'a> InlineSink<'a> for NotesSink<'_> {
    type Space = ();

    fn visit(&mut self) -> Result<(), String> {
        Ok(())
    }

    fn step(&mut self, _units: usize) -> Result<(), String> {
        Ok(())
    }

    fn enter(&mut self, _element: Nd<'a>) {}

    fn leave(&mut self) {}

    fn text(&mut self, piece: &str) -> Result<(), String> {
        self.push(piece)
    }

    fn hold_space(&mut self, _after_text: bool) -> Result<(), String> {
        Ok(())
    }

    fn put_space(&mut self, _space: ()) -> Result<(), String> {
        self.push(" ")
    }

    fn preserved(&mut self, class: Inline) -> Result<(), String> {
        match class {
            Inline::Space(count) => self.push(&" ".repeat(count)),
            Inline::Tab => self.push("\t"),
            _ => self.push("\n"),
        }
    }

    fn generated(&mut self, _field: Nd<'a>) -> Result<bool, String> {
        Ok(false)
    }
}

/// §6.1.2 white-space processing: runs of spaces, tabs and newlines in text
/// content collapse to one space (`text:s` carries significant spaces).
fn collapse_whitespace(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut last_space = false;
    for c in text.chars() {
        if matches!(c, ' ' | '\t' | '\n' | '\r') {
            if !last_space {
                out.push(' ');
            }
            last_space = true;
        } else {
            out.push(c);
            last_space = false;
        }
    }
    out
}

fn push_text(
    runs: &mut Vec<TextRun>,
    style: &RunStyle,
    text: &str,
    link: Option<&str>,
    env: &TextEnv<'_, '_>,
) -> Result<(), String> {
    if text.is_empty() {
        return Ok(());
    }
    // Resolving the run's link reads (hashes) its whole target.
    env.payload.step(link.map_or(0, |l| l.len() / 64))?;
    // The text (a case transform may grow it; it is charged again with
    // each run record below) before any copy is made.
    env.payload.charge(text.len())?;
    let (hyperlink, action) = match link {
        Some(href) => match href.strip_prefix('#') {
            Some(page) => match env.pages.get(page) {
                Some(index) => (
                    Some(slide_part_name(*index)),
                    Some("ppaction://hlinksldjump".to_string()),
                ),
                None => (None, None),
            },
            None => (Some(href.to_string()), None),
        },
        None => (None, None),
    };
    let transformed = match style.transform {
        Some("lowercase") => text.to_lowercase(),
        Some("capitalize") => capitalize(text),
        _ => text.to_string(),
    };
    // Split into script portions when the Asian or complex values differ.
    let needs_split = style.size[1] != style.size[0]
        || style.size[2] != style.size[0]
        || style.bold[1] != style.bold[0]
        || style.bold[2] != style.bold[0]
        || style.italic[1] != style.italic[0]
        || style.italic[2] != style.italic[0];
    // Each portion becomes a run, charged (the record its style copy, link
    // and text will make) before the copy is made, and kept at once.
    let style_bytes = record_bytes(&style.data)?;
    let link_bytes =
        hyperlink.as_ref().map_or(0, String::len) + action.as_ref().map_or(0, String::len);
    let mut emit = |script: Script, portion: String| -> Result<(), String> {
        env.payload.charge(
            style_bytes
                .saturating_add(std::mem::size_of::<TextRun>())
                .saturating_add(link_bytes)
                .saturating_add(portion.len()),
        )?;
        let index = match script {
            Script::Western => 0,
            Script::Asian => 1,
            Script::Complex => 2,
        };
        let mut data = style.data.clone();
        data.text = portion;
        if index > 0 {
            data.font_size = style.size[index].or(style.size[0]);
            data.bold = style.bold[index].or(style.bold[0]);
            data.italic = style.italic[index].or(style.italic[0]);
        }
        data.hyperlink = hyperlink.clone();
        data.hyperlink_action = action.clone();
        data.hyperlink_uses_text_fill = hyperlink.is_some();
        runs.push(TextRun::Text(data));
        Ok(())
    };
    if needs_split {
        let mut current = Script::Western;
        let mut buffer = String::new();
        let mut started = false;
        for c in transformed.chars() {
            let script = script_of(c).unwrap_or(current);
            if started && script != current && !buffer.is_empty() {
                emit(current, std::mem::take(&mut buffer))?;
            }
            current = script;
            started = true;
            buffer.push(c);
        }
        if !buffer.is_empty() {
            emit(current, buffer)?;
        }
    } else {
        emit(Script::Western, transformed)?;
    }
    Ok(())
}

fn capitalize(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut start = true;
    for c in text.chars() {
        if start && c.is_alphabetic() {
            out.extend(c.to_uppercase());
            start = false;
        } else {
            out.push(c);
            if c.is_whitespace() {
                start = true;
            }
        }
    }
    out
}

pub fn slide_part_name(index: usize) -> String {
    format!("odp/slide{}", index + 1)
}

/// Plain text of speaker notes: the paragraphs of the containers, one per
/// line, within [`MAX_NOTES_BYTES`] (checked before each piece is added).
pub fn notes_text<'a>(containers: impl Iterator<Item = Nd<'a>>) -> Result<String, String> {
    let mut text = String::new();
    let mut first = true;
    for container in containers {
        // Each outermost paragraph, in document order, is one line read as
        // slide paragraphs are (a paragraph nested in another, not valid
        // ODF, is read as part of it).
        let mut stack: Vec<Nd<'a>> = container.children().collect();
        stack.reverse();
        while let Some(node) = stack.pop() {
            if xml::is(node, NS_TEXT, "p") || xml::is(node, NS_TEXT, "h") {
                if !first {
                    NotesSink { text: &mut text }.push("\n")?;
                }
                first = false;
                let mut sink = NotesSink { text: &mut text };
                walk_paragraph(node, &mut sink, &mut Collapse::default(), 0)?;
                continue;
            }
            stack.extend(node.children().collect::<Vec<_>>().into_iter().rev());
        }
    }
    Ok(text)
}

/// Text content of an element whose children are inline text (a
/// `presentation:footer-decl`, §9.10.1), within [`MAX_NOTES_BYTES`].
pub fn plain_text_inline(node: Nd<'_>) -> Result<Rc<str>, String> {
    let mut text = String::new();
    let mut push = |piece: &str| -> Result<(), String> {
        if text.len().saturating_add(piece.len()) > MAX_NOTES_BYTES {
            return Err("OpenDocument header or footer exceeds the text limit".to_string());
        }
        text.push_str(piece);
        Ok(())
    };
    for d in node.descendants() {
        if d.is_text() {
            push(&collapse_whitespace(d.text().unwrap_or("")))?;
        } else if xml::is(d, NS_TEXT, "s") {
            let n = d
                .attribute((NS_TEXT, "c"))
                .and_then(|c| c.parse::<usize>().ok())
                .unwrap_or(1)
                .min(1024);
            for _ in 0..n {
                push(" ")?;
            }
        } else if xml::is(d, NS_TEXT, "tab") {
            push("\t")?;
        }
    }
    Ok(Rc::from(text))
}
