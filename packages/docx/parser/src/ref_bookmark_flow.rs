//! Bounded acquisition of one REF bookmark structure that cached result runs
//! cannot express: a leading authored page break.

use std::collections::{HashMap, HashSet};

use ooxml_common::ns::is_w_ns;
use roxmltree::Node;

use crate::types::{BreakType, ComplexFieldBoundaryWire, DocParagraph, DocRun};
use crate::xml_util::attr_w;

const MAX_REF_TARGETS: usize = 1024;
const MAX_ACTIVE_BOOKMARKS: usize = 16;
const MAX_TARGET_TEXT_BYTES: usize = 4096;
const MAX_TOTAL_TARGET_TEXT_BYTES: usize = 1024 * 1024;
const MAX_FIELD_INSTRUCTION_BYTES: usize = 256;

fn is_word(node: Node<'_, '_>, local_name: &str) -> bool {
    node.is_element()
        && is_w_ns(node.tag_name().namespace())
        && node.tag_name().name() == local_name
}

pub(crate) fn ref_target(instruction: &str) -> Option<&str> {
    let mut words = instruction.split_whitespace();
    words
        .next()
        .filter(|word| word.eq_ignore_ascii_case("REF"))?;
    // Quoted bookmark arguments need a full field-instruction tokenizer. A
    // split quoted name must never be mistaken for a different short bookmark.
    let target = words.next().filter(|word| {
        !word.is_empty() && !word.starts_with('\\') && !word.contains(['\'', '"'])
    })?;
    // §17.16.5.51 switches such as \p, \n, \r and \w change the result to
    // positional/numbering information. The four Office controls cover plain
    // REF and \h only; unknown switches keep the cached result untouched.
    match (words.next(), words.next()) {
        (None, None) => Some(target),
        (Some(switch), None) if switch.eq_ignore_ascii_case("\\h") => Some(target),
        _ => None,
    }
}

/// Collect only REF names required by this story. A field may cross projected
/// blocks, so the small instruction stack survives each block's XML arena.
#[derive(Default)]
pub(crate) struct RefInstructionCollector {
    stack: Vec<Option<String>>,
    overflow_depth: usize,
    pub(crate) targets: HashSet<String>,
}

impl RefInstructionCollector {
    pub(crate) fn observe(&mut self, root: Node<'_, '_>) {
        for node in root.descendants().filter(Node::is_element) {
            self.observe_node(node);
        }
    }

    pub(crate) fn observe_node(&mut self, node: Node<'_, '_>) {
        if is_word(node, "fldChar") {
            match attr_w(node, "fldCharType").as_deref() {
                Some("begin") => {
                    if self.overflow_depth > 0 || self.stack.len() >= MAX_ACTIVE_BOOKMARKS {
                        self.overflow_depth = self.overflow_depth.saturating_add(1);
                    } else {
                        self.stack.push(Some(String::new()));
                    }
                }
                Some("separate") if self.overflow_depth == 0 => {
                    if let Some(Some(instruction)) = self.stack.last() {
                        if let Some(target) = ref_target(instruction) {
                            if self.targets.len() < MAX_REF_TARGETS {
                                self.targets.insert(target.to_string());
                            }
                        }
                    }
                }
                Some("end") => {
                    if self.overflow_depth > 0 {
                        self.overflow_depth -= 1;
                    } else {
                        self.stack.pop();
                    }
                }
                _ => {}
            }
        } else if is_word(node, "instrText") && self.overflow_depth == 0 {
            if let Some(frame) = self.stack.last_mut() {
                if let Some(instruction) = frame {
                    let text = node.text().unwrap_or_default();
                    if instruction.len().saturating_add(text.len()) <= MAX_FIELD_INSTRUCTION_BYTES {
                        instruction.push_str(text);
                    } else {
                        *frame = None;
                    }
                }
            }
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Phase {
    BeforeBreak,
    AfterBreak,
    Text,
    Invalid,
}

struct ActiveBookmark {
    id: String,
    name: String,
    phase: Phase,
    text: String,
}

/// ECMA-376 Part 1 §17.16.5.51 defines REF as the bookmarked text or graphics.
/// Word controls with/without `\\h` and with/without a bookmarked leading
/// `<w:br w:type="page"/>` show that this break also precedes the field result.
/// We retain only a strict, bounded subset: one initial break paragraph and
/// one following text-only paragraph. Other structures, stale cached results,
/// nested targeted bookmarks, and budget overflows use the cached field result.
pub(crate) struct LeadingBreakCollector<'a> {
    requested: &'a HashSet<String>,
    active: Vec<ActiveBookmark>,
    completed: HashMap<String, String>,
    seen: HashSet<String>,
    ambiguous: HashSet<String>,
    total_text_bytes: usize,
}

impl<'a> LeadingBreakCollector<'a> {
    pub(crate) fn new(requested: &'a HashSet<String>) -> Self {
        Self {
            requested,
            active: Vec::new(),
            completed: HashMap::new(),
            seen: HashSet::new(),
            ambiguous: HashSet::new(),
            total_text_bytes: 0,
        }
    }

    pub(crate) fn observe(&mut self, root: Node<'_, '_>) {
        if !is_word(root, "p") {
            self.observe_other_block();
            return;
        }
        if self.requested.is_empty() {
            return;
        }
        if self.active.is_empty()
            && !root.children().any(|child| {
                is_word(child, "bookmarkStart")
                    && attr_w(child, "name").is_some_and(|name| self.requested.contains(&name))
            })
        {
            return;
        }
        let extra_structural_props = root
            .children()
            .find(|child| is_word(*child, "pPr"))
            .is_some_and(|properties| {
                properties
                    .descendants()
                    .any(|child| is_word(child, "pageBreakBefore") || is_word(child, "sectPr"))
            });
        for bookmark in &mut self.active {
            bookmark.phase = match bookmark.phase {
                Phase::AfterBreak if !extra_structural_props => Phase::Text,
                _ => Phase::Invalid,
            };
        }

        for node in root.descendants().filter(Node::is_element) {
            if is_word(node, "bookmarkStart") {
                if !node.parent().is_some_and(|parent| is_word(parent, "p")) {
                    continue;
                }
                let Some(name) = attr_w(node, "name") else {
                    continue;
                };
                if !self.requested.contains(&name) {
                    continue;
                }
                // A duplicate or overlapping requested range is ambiguous.
                if !self.seen.insert(name.clone()) {
                    self.completed.remove(&name);
                    self.ambiguous.insert(name);
                    continue;
                }
                if self.active.len() >= MAX_ACTIVE_BOOKMARKS || !self.active.is_empty() {
                    for bookmark in &mut self.active {
                        bookmark.phase = Phase::Invalid;
                    }
                    continue;
                }
                if let Some(id) = attr_w(node, "id") {
                    self.active.push(ActiveBookmark {
                        id,
                        name,
                        phase: if extra_structural_props {
                            Phase::Invalid
                        } else {
                            Phase::BeforeBreak
                        },
                        text: String::new(),
                    });
                }
            } else if is_word(node, "bookmarkEnd") {
                let Some(id) = attr_w(node, "id") else {
                    continue;
                };
                if let Some(index) = self.active.iter().position(|bookmark| bookmark.id == id) {
                    let bookmark = self.active.remove(index);
                    if bookmark.phase == Phase::Text
                        && !bookmark.text.is_empty()
                        && !self.ambiguous.contains(&bookmark.name)
                    {
                        self.total_text_bytes =
                            self.total_text_bytes.saturating_add(bookmark.text.len());
                        if self.total_text_bytes <= MAX_TOTAL_TARGET_TEXT_BYTES {
                            self.completed.insert(bookmark.name, bookmark.text);
                        }
                    }
                }
            } else if is_word(node, "br") {
                let page = attr_w(node, "type").as_deref() == Some("page");
                for bookmark in &mut self.active {
                    bookmark.phase = if page && bookmark.phase == Phase::BeforeBreak {
                        Phase::AfterBreak
                    } else {
                        Phase::Invalid
                    };
                }
            } else if is_word(node, "t") {
                let text = node.text().unwrap_or_default();
                for bookmark in &mut self.active {
                    if bookmark.phase == Phase::Text
                        && bookmark.text.len().saturating_add(text.len()) <= MAX_TARGET_TEXT_BYTES
                    {
                        bookmark.text.push_str(text);
                    } else {
                        bookmark.phase = Phase::Invalid;
                    }
                }
            } else if is_word(node, "r") {
                // Run formatting does not change the structural break. Any
                // non-text material would require full REF range expansion.
                let plain = node.children().filter(Node::is_element).all(|child| {
                    is_word(child, "rPr")
                        || is_word(child, "t")
                        || is_word(child, "br")
                        || is_word(child, "lastRenderedPageBreak")
                });
                if !plain {
                    for bookmark in &mut self.active {
                        bookmark.phase = Phase::Invalid;
                    }
                }
            } else if is_word(node, "hyperlink")
                || is_word(node, "drawing")
                || is_word(node, "fldSimple")
                || is_word(node, "ins")
                || is_word(node, "del")
            {
                for bookmark in &mut self.active {
                    bookmark.phase = Phase::Invalid;
                }
            }
        }
    }

    pub(crate) fn observe_other_block(&mut self) {
        // A table or other block in the bookmarked range is not the supported
        // single plain-text paragraph; its contents must not be skipped merely
        // because the projector emits it as a separate top-level block.
        for bookmark in &mut self.active {
            bookmark.phase = Phase::Invalid;
        }
    }

    pub(crate) fn finish(self) -> HashMap<String, String> {
        self.completed
    }
}

/// Insert the one structurally proven break before the cached result. Field
/// boundary indices move with the inserted run; the normal paragraph splitter
/// turns it into a body page break. No cached text is regenerated.
pub(crate) fn apply_matching_leading_break(
    paragraph: &mut DocParagraph,
    targets: &HashMap<String, String>,
) {
    if targets.is_empty() {
        return;
    }
    let mut starts: HashMap<u32, (usize, &ComplexFieldBoundaryWire)> = HashMap::new();
    let mut insertions = Vec::new();
    for (boundary_ordinal, boundary) in paragraph.complex_field_boundaries.iter().enumerate() {
        if boundary.boundary == "start" && boundary.field_type == "ref" {
            starts.insert(boundary.occurrence_id, (boundary_ordinal, boundary));
        } else if boundary.boundary == "end" {
            let Some((start_ordinal, begin)) = starts.remove(&boundary.occurrence_id) else {
                continue;
            };
            if begin.run_index >= boundary.run_index || boundary.run_index > paragraph.runs.len() {
                continue;
            }
            // A nested field has its own semantics; a text-only cached result
            // does not prove where its bookmark structure belongs. Adjacent
            // boundary events are the strict non-nested case, checked in O(1).
            if boundary_ordinal != start_ordinal + 1 {
                continue;
            }
            let Some(target) = ref_target(&begin.instruction).and_then(|name| targets.get(name))
            else {
                continue;
            };
            let mut cached = String::new();
            let mut plain = true;
            for run in &paragraph.runs[begin.run_index..boundary.run_index] {
                if let DocRun::Text(text) = run {
                    if cached.len().saturating_add(text.text.len()) > MAX_TARGET_TEXT_BYTES {
                        plain = false;
                        break;
                    }
                    cached.push_str(&text.text);
                } else {
                    plain = false;
                    break;
                }
            }
            if plain && cached == *target {
                insertions.push(begin.run_index);
            }
        }
    }
    insertions.sort_unstable();
    insertions.dedup();
    if insertions.is_empty() {
        return;
    }
    // Move each run once; repeated Vec::insert would become quadratic for a
    // paragraph containing many independent REF fields.
    let old_runs = std::mem::take(&mut paragraph.runs);
    let old_len = old_runs.len();
    let mut rebuilt = Vec::with_capacity(old_len.saturating_add(insertions.len()));
    let mut next_insertion = 0;
    for (index, run) in old_runs.into_iter().enumerate() {
        if insertions.get(next_insertion) == Some(&index) {
            rebuilt.push(DocRun::Break {
                break_type: BreakType::Page,
            });
            next_insertion += 1;
        }
        rebuilt.push(run);
    }
    paragraph.runs = rebuilt;
    // Boundary indices are normally in story order, but a binary partition
    // keeps this correct even if a caller supplies them in another order.
    for boundary in &mut paragraph.complex_field_boundaries {
        let old_index = boundary.run_index;
        boundary.run_index =
            old_index.saturating_add(insertions.partition_point(|index| *index <= old_index));
    }
}
