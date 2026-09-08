use crate::styles::{parse_run_fmt, RunFmt};
use crate::xml_util::*;
use ooxml_common::blip::mime_from_ext;
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::{attr_ns, relationships};
#[cfg(test)]
use ooxml_common::numbering::format_counter;
use ooxml_common::numbering::{CounterEngine, CounterError, CounterIdentity, LevelFacts};
use std::cell::Cell;
use std::collections::HashMap;
use std::rc::Rc;

#[cfg(test)]
#[path = "numbering/restart_tests.rs"]
mod restart_tests;

#[cfg(test)]
#[path = "numbering/legal_tests.rs"]
mod legal_tests;

#[cfg(test)]
#[path = "numbering/counter_instance_tests.rs"]
mod counter_instance_tests;

/// Parse a single VML CSS length (e.g. `width:9pt`) from a `style` attribute
/// into pt. Supports the units Word emits for picture-bullet shapes: `pt`
/// (1pt), `in` (72pt), `pc`/`pi` (12pt), `cm` (28.3465pt), `mm` (2.83465pt). A
/// bare number with no unit is treated as pt (VML's default user unit for
/// shapes is the point). Returns `None` when the property is absent or
/// unparseable.
fn vml_style_len(style: &str, prop: &str) -> Option<f64> {
    for decl in style.split(';') {
        let (k, v) = decl.split_once(':')?;
        if k.trim() != prop {
            continue;
        }
        return parse_measure_to_pt(v, 1.0);
    }
    None
}

#[derive(Debug, Clone)]
pub struct LevelDef {
    pub format: String,   // "decimal" | "bullet" | etc.
    pub text: String,     // lvlText val, e.g. "%1." or "•"
    pub indent_left: f64, // pt — w:ind@left ≡ logical START indent (§17.3.1.12)
    /// Whether `indent_left` came from this level's authored `w:ind@left`.
    /// False means the value is only the legacy depth fallback used when no
    /// lower style layer supplies an indent.
    pub indent_left_authored: bool,
    /// pt — w:ind@right ≡ logical END indent (Part 4 §14.11.2). RTL list levels
    /// carry their indent here (the renderer maps it to the physical left side).
    pub indent_right: Option<f64>,
    /// pt — SIGNED first-line indent (§17.3.1.12): `w:hanging` ⇒ negative (the
    /// marker hangs left of the body), `w:firstLine` ⇒ positive (an additional
    /// first-line indent). Mirrors how `styles.rs` stores a direct `w:ind`.
    pub indent_first: f64,
    /// Whether `indent_first` came from authored `w:hanging` / `w:firstLine`.
    pub indent_first_authored: bool,
    /// pt — POSITIVE magnitude of the first-line indent (= `indent_first.abs()`).
    /// Distinct from `indent_first` because the renderer uses it as the marker's
    /// tab-advance distance (§17.9.6 + §17.3.1.38, suff=tab), which is unsigned.
    pub tab: f64,
    /// ECMA-376 §17.9.28 `<w:suff>` — what follows the number text: "tab"
    /// (default), "space", or "nothing". Controls where the body text starts
    /// relative to the marker on the first line.
    pub suff: String,
    /// ECMA-376 §17.9.8 `<w:lvlJc>` — marker justification at its reference
    /// position: "left" (default, marker LEFT edge at the hanging-indent
    /// position), "right" (marker RIGHT edge there — period-aligned roman/decimal
    /// numerals), or "center". `<w:start>` is unrelated.
    pub lvl_jc: String,
    pub start: u32,
    /// ECMA-376 17.9.10: one-based last ancestor that resets this level.
    /// Zero means never; absence (or an invalid index) uses the previous level.
    restart: Option<u32>,
    /// ECMA-376 17.9.4: use decimal for every placeholder in this level's
    /// marker, without changing the referenced levels' own number formats.
    legal: bool,
    /// ECMA-376 §17.9.6 `<w:lvl><w:rPr>` — the level's run (character) properties
    /// for the number/bullet glyph itself. Merged OVER the paragraph's resolved
    /// run formatting at use-site so the marker's font axes (ascii/eastAsia)
    /// resolve through the same chain a body run uses. Often only carries a bare
    /// `<w:rFonts w:hint="eastAsia"/>` (no explicit typeface), in which case every
    /// axis is `None` and the marker simply inherits the paragraph's fonts.
    pub rpr: RunFmt,
    /// ECMA-376 §17.9.9 `<w:lvlPicBulletId w:val="N"/>` — when present, the
    /// level's marker is the image defined by the `<w:numPicBullet>` whose
    /// `numPicBulletId` is N (§17.9.20), drawn in place of `text`. Resolved at
    /// parse time to the bullet image's zip path (+ MIME + pt size from the
    /// `<v:shape style>`). `None` ⇒ ordinary text/glyph marker.
    pub pic_bullet: Option<PicBullet>,
    /// ECMA-376 §17.9.23 `<w:lvl><w:pStyle w:val>` — the styleId of the
    /// paragraph style ASSOCIATED with this level. Paragraphs of that style
    /// "shall automatically utilize this numbering level", and the ilvl carried
    /// by the STYLE's own `numPr` "shall be ignored" in its favor. Resolved via
    /// [`NumberingMap::level_for_style`] into each style's numbering level after
    /// both parts are parsed (see `StyleMap::resolve_numbering_level_backlinks`).
    /// `None` ⇒ no style association on this level.
    pub p_style: Option<String>,
}

/// ECMA-376 §17.9.20 `<w:numPicBullet>` — an image used as a list marker. The
/// image is defined by a VML `<w:pict><v:shape><v:imagedata r:id="…"/>` whose
/// `r:id` resolves through `word/_rels/numbering.xml.rels` to a media part, and
/// whose `<v:shape style="width:..;height:..">` carries the marker size.
#[derive(Debug, Clone)]
pub struct PicBullet {
    /// Zip path of the bullet image (e.g. `word/media/image1.gif`), resolved
    /// from the `<v:imagedata r:id>` via the numbering part's relationships.
    pub image_path: String,
    /// MIME type derived from the part extension (e.g. `image/gif`).
    pub mime_type: String,
    /// Marker width in pt, from the `<v:shape style="width:..">`. `None` when the
    /// shape style omits a width — ECMA-376 §17.9.20 derives the picture-bullet
    /// size from the drawing's own extent and defines no fallback dimension, so we
    /// surface the absence and let the renderer fall back to the resolved marker
    /// font size (its single source of truth) rather than inventing a magic pt.
    pub width_pt: Option<f64>,
    /// Marker height in pt, from the `<v:shape style="height:..">`. `None` ⇒ see
    /// {@link PicBullet::width_pt}.
    pub height_pt: Option<f64>,
}

impl Default for LevelDef {
    fn default() -> Self {
        LevelDef {
            format: "decimal".to_string(),
            text: "%1.".to_string(),
            indent_left: 36.0,
            indent_left_authored: false,
            indent_right: None,
            indent_first: -36.0,
            indent_first_authored: false,
            tab: 36.0,
            suff: "tab".to_string(),
            lvl_jc: "left".to_string(),
            start: 1,
            restart: None,
            legal: false,
            rpr: RunFmt::default(),
            pic_bullet: None,
            p_style: None,
        }
    }
}

#[derive(Default, Clone)]
pub struct NumberingMap {
    /// abstractNumId → [level0..level8]
    abstract_nums: HashMap<u32, Vec<LevelDef>>,
    /// numId → abstractNumId
    num_to_abstract: HashMap<u32, u32>,
    /// numId → level override starts
    num_overrides: HashMap<u32, HashMap<u32, u32>>,
    /// ECMA-376 §17.9.7 — numId → per-level FULL `<w:lvl>` replacements from
    /// `<w:num><w:lvlOverride><w:lvl>`: "the numbering level formatting which
    /// shall be substituted for the given numbering level of the abstract
    /// definition". A replacement alone does not immediately restart a live
    /// counter (`startOverride`, tracked in `num_overrides`, does that). Its
    /// `lvlRestart` still controls later ancestor-triggered resets. Consulted
    /// before the abstract's levels in `get_level`,
    /// so lvlText/numFmt/indents/rPr/pStyle all substitute per-numId.
    num_level_overrides: HashMap<u32, HashMap<u32, LevelDef>>,
    /// Per-**abstractNumId** per-level counter. ECMA-376 §17.9.1 and §17.9.15
    /// define abstract numbering definitions and concrete numbering instances;
    /// they do not by themselves establish this runtime counter-store policy.
    /// The §17.9.26 `startOverride` example establishes shared restart behavior:
    /// numIds 5, 5, 6, 5 on one abstractNum produce 1, 2, 1, 2. The wider alias
    /// and full-level-replacement policy is bounded native-Word compatibility
    /// evidence (see `counter_instance_tests`). The shared engine uses disjoint
    /// `CounterIdentity` variants so an unresolved numId gets a disjoint counter
    /// instead of colliding with an abstractNumId. It also owns the (numId,
    /// level) pairs already advanced at
    /// least once. A numId carrying a
    /// `<w:lvlOverride><w:startOverride>` restarts the shared abstract counter
    /// only on its FIRST appearance at that level. ECMA-376 §17.9.26 defines
    /// `startOverride` and demonstrates its reset propagating across numIds on
    /// one abstractNum; applying it only once per numId follows measured
    /// native-Word compatibility behavior.
    counters: CounterEngine<u32, u32, u32>,
    counter_error: Rc<Cell<Option<CounterError>>>,
}

/// Parse one `<w:lvl>` element (ECMA-376 §17.9.6) into a [`LevelDef`].
///
/// Shared by the two places a level definition may appear: inside
/// `<w:abstractNum>` (§17.9.1) and as the FULL replacement inside a
/// `<w:num><w:lvlOverride>` (§17.9.7 — "the numbering level formatting which
/// shall be substituted for the given numbering level"). `depth` is the level's
/// 0-based position, used only for the spec-less fallback indent when the level
/// carries no `w:ind` at all.
fn parse_level_def(
    lvl_node: roxmltree::Node,
    depth: usize,
    pic_bullets: &HashMap<u32, PicBullet>,
) -> LevelDef {
    let start = child_w(lvl_node, "start")
        .and_then(|n| attr_w(n, "val"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    // ECMA-376 17.9.10 also applies to complete level replacements (17.9.5).
    // MS-OE376 2.1.285(b) records Word ignoring this property in replacements;
    // keep the normative OOXML behavior instead of adding an Office heuristic.
    let restart = child_w(lvl_node, "lvlRestart")
        .and_then(|n| attr_w(n, "val"))
        .and_then(|v| v.parse::<u32>().ok());
    let legal = bool_prop(lvl_node, "isLgl").unwrap_or(false);
    let format = child_w(lvl_node, "numFmt")
        .and_then(|n| attr_w(n, "val"))
        .unwrap_or_else(|| "decimal".to_string());
    let text = child_w(lvl_node, "lvlText")
        .and_then(|n| attr_w(n, "val"))
        .unwrap_or_else(|| "%1.".to_string());
    let ind_node = child_w(lvl_node, "pPr").and_then(|p| child_w(p, "ind"));
    // When the level defines a w:ind, a missing @left means "no
    // start indent from this source" (an RTL level carries its
    // indent in @right ≡ end instead); the per-level depth default
    // applies only when no w:ind exists at all.
    let indent_left_authored = ind_node.and_then(|i| attr_w(i, "left")).is_some();
    let indent_left = ind_node
        .and_then(|i| attr_w(i, "left"))
        .map(|v| twips_to_pt(&v))
        .unwrap_or(if ind_node.is_some() {
            0.0
        } else {
            720.0 / 20.0 * (depth as f64 + 1.0)
        });
    let indent_right = ind_node
        .and_then(|i| attr_w(i, "right"))
        .map(|v| twips_to_pt(&v));
    // §17.3.1.12: a first-line indent is EITHER `w:hanging` (negative —
    // the marker hangs left of the body) OR `w:firstLine` (positive — an
    // additional first-line indent); `hanging` wins when both appear, the
    // same precedence `styles.rs` applies to a direct `w:ind`. Keep the
    // SIGN in `indent_first`; `tab` keeps the positive magnitude for the
    // marker's tab-advance.
    let indent_first_authored = ind_node
        .is_some_and(|i| attr_w(i, "hanging").is_some() || attr_w(i, "firstLine").is_some());
    let indent_first = ind_node
        .and_then(|i| {
            attr_w(i, "hanging")
                .map(|v| -twips_to_pt(&v))
                .or_else(|| attr_w(i, "firstLine").map(|v| twips_to_pt(&v)))
        })
        .unwrap_or(-36.0);
    let tab = indent_first.abs();
    // §17.9.28: absent <w:suff> means "tab".
    let suff = child_w(lvl_node, "suff")
        .and_then(|n| attr_w(n, "val"))
        .unwrap_or_else(|| "tab".to_string());
    // §17.9.8 `<w:lvlJc>` — marker justification; absent ⇒ "left".
    let lvl_jc = child_w(lvl_node, "lvlJc")
        .and_then(|n| attr_w(n, "val"))
        .unwrap_or_else(|| "left".to_string());
    // §17.9.6 — the level's run properties for the marker glyph.
    // Parsed with the SAME `parse_run_fmt` body runs use; theme refs
    // stay as "@theme:<ref>" markers and are resolved at use-site once
    // merged over the paragraph's run formatting.
    let rpr = child_w(lvl_node, "rPr")
        .map(parse_run_fmt)
        .unwrap_or_default();
    // §17.9.9 — resolve the level's picture bullet (if any) against
    // the `<w:numPicBullet>` definitions collected by the caller.
    let pic_bullet = child_w(lvl_node, "lvlPicBulletId")
        .and_then(|n| attr_w(n, "val"))
        .and_then(|v| v.parse::<u32>().ok())
        .and_then(|id| pic_bullets.get(&id).cloned());
    // §17.9.23 — the paragraph style this level is associated with.
    let p_style = child_w(lvl_node, "pStyle").and_then(|n| attr_w(n, "val"));
    LevelDef {
        format,
        text,
        indent_left,
        indent_left_authored,
        indent_right,
        indent_first,
        indent_first_authored,
        tab,
        suff,
        lvl_jc,
        start,
        restart,
        legal,
        rpr,
        pic_bullet,
        p_style,
    }
}

impl NumberingMap {
    /// Parse `word/numbering.xml`. `media_map` is the numbering part's own
    /// relationship table (rId → zip media path, built from
    /// `word/_rels/numbering.xml.rels`); it is required to resolve the
    /// `<w:numPicBullet>` images (§17.9.20) — an empty map simply yields no
    /// picture bullets, leaving levels on their text/glyph markers.
    pub fn parse(xml: &str, media_map: &HashMap<String, String>) -> Self {
        let mut map = NumberingMap::default();
        let doc = match parse_guarded(xml) {
            Ok(d) => d,
            Err(_) => return map,
        };
        let root = doc.root_element();

        // ECMA-376 §17.9.20 — collect `<w:numPicBullet>` definitions first so
        // each level's `<w:lvlPicBulletId>` (§17.9.9) can resolve against them.
        // The bullet image is a VML `<v:shape><v:imagedata r:id>`; the r:id maps
        // to a media part through the numbering part's rels (`media_map`), and
        // the `<v:shape style="width:..;height:..">` carries the marker size.
        let mut pic_bullets: HashMap<u32, PicBullet> = HashMap::new();
        for pb_node in children_w(root, "numPicBullet") {
            let Some(id) = attr_w(pb_node, "numPicBulletId").and_then(|v| v.parse::<u32>().ok())
            else {
                continue;
            };
            let Some(imagedata) = pb_node
                .descendants()
                .find(|n| n.tag_name().name() == "imagedata")
            else {
                continue;
            };
            // `r:id` lives in the relationships namespace (Transitional or
            // Strict); fall back to the unqualified attribute for defensiveness.
            let Some(rid) = attr_ns(
                &imagedata,
                relationships::TRANSITIONAL,
                relationships::STRICT,
                "id",
            ) else {
                continue;
            };
            let Some(image_path) = media_map.get(rid).cloned() else {
                continue;
            };
            // `<v:shape style="width:9pt;height:9pt">` — VML CSS lengths. The
            // dimension is left as `None` when the style omits it: §17.9.20 has no
            // default picture-bullet size, so the renderer (not the parser)
            // resolves the absence against the marker font size.
            let shape_style = pb_node
                .descendants()
                .find(|n| n.tag_name().name() == "shape")
                .and_then(|n| n.attribute("style"))
                .unwrap_or("");
            let width_pt = vml_style_len(shape_style, "width");
            let height_pt = vml_style_len(shape_style, "height");
            let mime_type = mime_from_ext(&image_path).to_string();
            pic_bullets.insert(
                id,
                PicBullet {
                    image_path,
                    mime_type,
                    width_pt,
                    height_pt,
                },
            );
        }

        // Parse abstractNum definitions
        for abs_node in children_w(root, "abstractNum") {
            let Some(abs_id_s) = attr_w(abs_node, "abstractNumId") else {
                continue;
            };
            let abs_id: u32 = abs_id_s.parse().unwrap_or(0);
            let mut levels = vec![];
            for lvl_node in children_w(abs_node, "lvl") {
                let depth = levels.len();
                levels.push(parse_level_def(lvl_node, depth, &pic_bullets));
            }
            map.abstract_nums.insert(abs_id, levels);
        }

        // Parse num → abstractNum
        for num_node in children_w(root, "num") {
            let Some(num_id_s) = attr_w(num_node, "numId") else {
                continue;
            };
            let num_id: u32 = num_id_s.parse().unwrap_or(0);
            if let Some(abs_ref) = child_w(num_node, "abstractNumId").and_then(|n| attr_w(n, "val"))
            {
                let abs_id: u32 = abs_ref.parse().unwrap_or(0);
                map.num_to_abstract.insert(num_id, abs_id);
            }
            // Level overrides
            let mut overrides = HashMap::new();
            let mut level_overrides = HashMap::new();
            for lvl_ov in children_w(num_node, "lvlOverride") {
                let ilvl: u32 = attr_w(lvl_ov, "ilvl")
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                if let Some(start_ov) =
                    child_w(lvl_ov, "startOverride").and_then(|n| attr_w(n, "val"))
                {
                    overrides.insert(ilvl, start_ov.parse().unwrap_or(1));
                }
                // §17.9.7 — a FULL <w:lvl> child substitutes the level's
                // definition for this numId, without an immediate restart.
                if let Some(lvl_node) = child_w(lvl_ov, "lvl") {
                    level_overrides
                        .insert(ilvl, parse_level_def(lvl_node, ilvl as usize, &pic_bullets));
                }
            }
            if !overrides.is_empty() {
                map.num_overrides.insert(num_id, overrides);
            }
            if !level_overrides.is_empty() {
                map.num_level_overrides.insert(num_id, level_overrides);
            }
        }

        map
    }

    /// The EFFECTIVE level definition for (numId, level): the numId's own
    /// `<w:lvlOverride><w:lvl>` substitution when present (§17.9.7), else the
    /// abstract definition's level (§17.9.6).
    pub fn get_level(&self, num_id: u32, level: u32) -> Option<&LevelDef> {
        if let Some(ov) = self
            .num_level_overrides
            .get(&num_id)
            .and_then(|m| m.get(&level))
        {
            return Some(ov);
        }
        let abs_id = self.num_to_abstract.get(&num_id)?;
        let levels = self.abstract_nums.get(abs_id)?;
        levels.get(level as usize)
    }

    /// ECMA-376 §17.9.23 — the numbering level ASSOCIATED with a paragraph
    /// style inside the numbering definition that `num_id` references: the
    /// first level whose `<w:pStyle w:val>` equals `style_id`. Paragraphs of
    /// that style "shall automatically utilize this numbering level"; the ilvl
    /// in the STYLE's own `numPr` "shall be ignored" in its favor. Goes through
    /// [`Self::get_level`] so a backlink carried by a per-numId `<w:lvlOverride>`
    /// substitution (§17.9.7) participates too. `None` ⇒ the list has no
    /// association for this style (or the numId dangles). WordprocessingML caps
    /// This model supports nine levels, matching CT_AbstractNum's maximum of
    /// nine lvl children (§A.1); ilvl itself uses ST_DecimalNumber (§17.9.3).
    pub fn level_for_style(&self, num_id: u32, style_id: &str) -> Option<u32> {
        (0..9).find(|&l| {
            self.get_level(num_id, l)
                .is_some_and(|def| def.p_style.as_deref() == Some(style_id))
        })
    }

    /// Advance the counter for (numId, level), resetting deeper levels.
    ///
    /// The counter is keyed by the numId's **abstractNumId**, so all numIds that
    /// share an abstract definition advance one running count. This storage
    /// policy follows measured native-Word behavior; see the `counters` field
    /// doc and `counter_instance_tests`. Each level stores its CURRENT value (not
    /// the next): a level's first appearance shows its `start`, each later
    /// advance adds one. Advancing a level resets descendants whose effective
    /// `lvlRestart` includes that ancestor (17.9.10); by default this means all
    /// deeper levels. Shallower levels are seeded to their `start` so an
    /// ancestor that only prefixes the marker (e.g. `%1.%2`) still resolves when
    /// it is never advanced on its own.
    ///
    /// A numId whose `<w:lvlOverride>` carries a `<w:startOverride>` for this
    /// level RESTARTS the shared abstract counter to the override value on its
    /// first appearance at that level, then increments normally. ECMA-376
    /// §17.9.26 defines `startOverride` and demonstrates the shared reset with
    /// numIds 5, 5, 6, 5 producing 1, 2, 1, 2. The once-per-numId application is
    /// measured native-Word compatibility behavior. Returns the value to display.
    pub fn advance(&mut self, num_id: u32, level: u32) -> Result<u32, CounterError> {
        let identity = counter_identity(&self.num_to_abstract, num_id);
        let abstract_nums = &self.abstract_nums;
        let level_overrides = &self.num_level_overrides;
        let start_override = self
            .num_overrides
            .get(&num_id)
            .and_then(|levels| levels.get(&level))
            .copied();
        self.counters.advance(
            identity,
            num_id,
            level,
            start_override,
            |index| {
                self.num_overrides
                    .get(&num_id)
                    .and_then(|levels| levels.get(&index))
                    .copied()
                    .or_else(|| {
                        level_facts(
                            abstract_nums,
                            level_overrides,
                            &self.num_to_abstract,
                            num_id,
                            index,
                        )
                        .map(|facts| facts.start)
                    })
                    .unwrap_or(1)
            },
            |index| {
                level_facts(
                    abstract_nums,
                    level_overrides,
                    &self.num_to_abstract,
                    num_id,
                    index,
                )
            },
        )
    }

    /// Resolve the display text for a counter value in the given level.
    ///
    /// ECMA-376 §17.9.11 (`<w:lvlText>`): each `%N` placeholder is the counter
    /// of level `N-1`, formatted with THAT level's own `<w:numFmt>`. A
    /// multi-level marker such as `%1.%2` therefore needs every ancestor
    /// counter, not just the current level's. The current level uses `counter`
    /// (the value `advance` just returned); ancestor levels read their live
    /// counter from `self.counters` — `advance` seeds every shallower level to
    /// its start, so an ancestor that is never itself advanced (e.g. a list
    /// whose level 0 only exists to prefix subsection numbers with a fixed
    /// `start`) still resolves to its start value.
    pub fn resolve_text(
        &self,
        num_id: u32,
        level: u32,
        counter: u32,
    ) -> Result<String, CounterError> {
        let identity = counter_identity(&self.num_to_abstract, num_id);
        self.counters.resolve_text(
            identity,
            level,
            counter,
            |index| {
                self.num_overrides
                    .get(&num_id)
                    .and_then(|levels| levels.get(&index))
                    .copied()
                    .or_else(|| self.get_level(num_id, index).map(|facts| facts.start))
                    .unwrap_or(1)
            },
            |index| {
                level_facts(
                    &self.abstract_nums,
                    &self.num_level_overrides,
                    &self.num_to_abstract,
                    num_id,
                    index,
                )
            },
        )
    }

    pub fn record_counter_error(&self, error: CounterError) {
        if self.counter_error.get().is_none() {
            self.counter_error.set(Some(error));
        }
    }

    pub fn check_counter_error(&self) -> Result<(), String> {
        match self.counter_error.get() {
            None => Ok(()),
            Some(CounterError::InvalidLevel) => Err("unsupported numbering level".to_string()),
            Some(CounterError::Overflow) => Err("numbering counter overflow".to_string()),
            Some(CounterError::OutputTooLarge) => {
                Err("numbering marker output too large".to_string())
            }
        }
    }
}

fn counter_identity(aliases: &HashMap<u32, u32>, num_id: u32) -> CounterIdentity<u32, u32> {
    aliases
        .get(&num_id)
        .copied()
        .map_or(CounterIdentity::Orphan(num_id), CounterIdentity::Shared)
}

fn level_facts<'a>(
    abstracts: &'a HashMap<u32, Vec<LevelDef>>,
    overrides: &'a HashMap<u32, HashMap<u32, LevelDef>>,
    aliases: &HashMap<u32, u32>,
    num_id: u32,
    level: u32,
) -> Option<LevelFacts<'a>> {
    let definition = overrides
        .get(&num_id)
        .and_then(|levels| levels.get(&level))
        .or_else(|| {
            aliases
                .get(&num_id)
                .and_then(|id| abstracts.get(id))
                .and_then(|levels| levels.get(level as usize))
        })?;
    Some(LevelFacts {
        start: definition.start,
        restart: definition.restart,
        format: &definition.format,
        text: &definition.text,
        legal: definition.legal,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    fn map(body: &str) -> NumberingMap {
        NumberingMap::parse(
            &format!("<w:numbering {W}>{body}</w:numbering>"),
            &HashMap::new(),
        )
    }

    /// §17.9.20 / §17.9.9 — a `<w:numPicBullet>` image resolves through the
    /// numbering part's rels (`media_map`), and the level's `<w:lvlPicBulletId>`
    /// picks it up with the `<v:shape style>` size (here width:9pt;height:9pt).
    #[test]
    fn picture_bullet_resolves_image_path_and_size() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        const O: &str = "xmlns:o=\"urn:schemas-microsoft-com:office:office\"";
        let media: HashMap<String, String> =
            [("rId1".to_string(), "word/media/image1.gif".to_string())]
                .into_iter()
                .collect();
        let xml = format!(
            r#"<w:numbering {W} {R} {V} {O}>
                 <w:numPicBullet w:numPicBulletId="0">
                   <w:pict>
                     <v:shape id="x" style="width:9pt;height:9pt" o:bullet="t">
                       <v:imagedata r:id="rId1" o:title="BD"/>
                     </v:shape>
                   </w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="8">
                   <w:lvl w:ilvl="0">
                     <w:numFmt w:val="bullet"/><w:lvlText w:val=""/>
                     <w:lvlPicBulletId w:val="0"/>
                   </w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &media);
        let lvl = m.get_level(3, 0).expect("level 0");
        let pb = lvl.pic_bullet.as_ref().expect("picture bullet resolved");
        assert_eq!(pb.image_path, "word/media/image1.gif");
        assert_eq!(pb.mime_type, "image/gif");
        assert!((pb.width_pt.expect("width from style") - 9.0).abs() < 1e-6);
        assert!((pb.height_pt.expect("height from style") - 9.0).abs() < 1e-6);
    }

    /// §17.9.20 — when the `<v:shape>` style omits width/height, the size is left
    /// as `None` (no magic 9pt default in the parser). The renderer falls back to
    /// the resolved marker font size, so the parser must NOT invent a dimension.
    #[test]
    fn picture_bullet_without_shape_size_is_none() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        let media: HashMap<String, String> =
            [("rId1".to_string(), "word/media/image1.png".to_string())]
                .into_iter()
                .collect();
        let xml = format!(
            r#"<w:numbering {W} {R} {V}>
                 <w:numPicBullet w:numPicBulletId="0">
                   <w:pict><v:shape id="x">
                     <v:imagedata r:id="rId1"/>
                   </v:shape></w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="8">
                   <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>
                     <w:lvlPicBulletId w:val="0"/></w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &media);
        let pb = m
            .get_level(3, 0)
            .unwrap()
            .pic_bullet
            .as_ref()
            .expect("picture bullet resolved");
        assert_eq!(pb.image_path, "word/media/image1.png");
        assert_eq!(pb.mime_type, "image/png");
        assert_eq!(
            pb.width_pt, None,
            "no shape width ⇒ None (no magic default)"
        );
        assert_eq!(pb.height_pt, None);
    }

    /// §17.9.20 → §17.9.9 end-to-end resolution chain: a `<w:numPicBullet>` (id N)
    /// whose `<v:imagedata r:id>` resolves through the numbering part's rels, then
    /// a level's `<w:lvlPicBulletId w:val="N"/>` picks that bullet up. Confirms the
    /// id wiring (not just a single happy-path size): a DIFFERENT id is rejected
    /// and the matching id surfaces the right media path + size.
    #[test]
    fn lvl_pic_bullet_id_resolves_matching_num_pic_bullet() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        let media: HashMap<String, String> = [
            ("rId7".to_string(), "word/media/bullet-a.png".to_string()),
            ("rId8".to_string(), "word/media/bullet-b.gif".to_string()),
        ]
        .into_iter()
        .collect();
        let xml = format!(
            r#"<w:numbering {W} {R} {V}>
                 <w:numPicBullet w:numPicBulletId="1">
                   <w:pict><v:shape style="width:12pt;height:6pt">
                     <v:imagedata r:id="rId7"/></v:shape></w:pict>
                 </w:numPicBullet>
                 <w:numPicBullet w:numPicBulletId="2">
                   <w:pict><v:shape style="width:8pt;height:8pt">
                     <v:imagedata r:id="rId8"/></v:shape></w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="4">
                   <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>
                     <w:lvlPicBulletId w:val="2"/></w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="9"><w:abstractNumId w:val="4"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &media);
        let pb = m
            .get_level(9, 0)
            .unwrap()
            .pic_bullet
            .as_ref()
            .expect("lvlPicBulletId=2 resolves to numPicBulletId=2");
        // The level referenced id=2, so it must surface bullet-b (NOT bullet-a).
        assert_eq!(pb.image_path, "word/media/bullet-b.gif");
        assert_eq!(pb.mime_type, "image/gif");
        assert!((pb.width_pt.unwrap() - 8.0).abs() < 1e-6);
        assert!((pb.height_pt.unwrap() - 8.0).abs() < 1e-6);
    }

    /// An unresolvable `r:id` (no matching rel) yields no picture bullet — the
    /// level falls back to its ordinary text marker.
    #[test]
    fn picture_bullet_missing_rel_falls_back() {
        const R: &str =
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        const V: &str = "xmlns:v=\"urn:schemas-microsoft-com:vml\"";
        let xml = format!(
            r#"<w:numbering {W} {R} {V}>
                 <w:numPicBullet w:numPicBulletId="0">
                   <w:pict><v:shape style="width:9pt;height:9pt">
                     <v:imagedata r:id="rIdX"/>
                   </v:shape></w:pict>
                 </w:numPicBullet>
                 <w:abstractNum w:abstractNumId="8">
                   <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>
                     <w:lvlPicBulletId w:val="0"/></w:lvl>
                 </w:abstractNum>
                 <w:num w:numId="3"><w:abstractNumId w:val="8"/></w:num>
               </w:numbering>"#
        );
        let m = NumberingMap::parse(&xml, &HashMap::new());
        assert!(m.get_level(3, 0).unwrap().pic_bullet.is_none());
    }

    /// §17.9.11 — a subsection list (`%1.%2`) whose level 0 is never advanced
    /// but starts at 3 must render "3.1", "3.2", … (the bug: only the current
    /// level's placeholder was substituted, leaving a literal "%1").
    #[test]
    fn multilevel_parent_placeholder_uses_level_start_when_not_advanced() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="5">
                 <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>"#);
        let c1 = m.advance(5, 1).unwrap();
        assert_eq!(m.resolve_text(5, 1, c1).unwrap(), "3.1");
        let c2 = m.advance(5, 1).unwrap();
        assert_eq!(m.resolve_text(5, 1, c2).unwrap(), "3.2");
        let c3 = m.advance(5, 1).unwrap();
        assert_eq!(m.resolve_text(5, 1, c3).unwrap(), "3.3");
    }

    /// Parent counter is tracked live and resets deeper levels: 1, 1.1, 1.2,
    /// 2, 2.1.
    #[test]
    fn multilevel_parent_counter_increments_and_resets() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="0">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#);
        let a = m.advance(1, 0).unwrap();
        assert_eq!(m.resolve_text(1, 0, a).unwrap(), "1.");
        let b = m.advance(1, 1).unwrap();
        assert_eq!(m.resolve_text(1, 1, b).unwrap(), "1.1");
        let c = m.advance(1, 1).unwrap();
        assert_eq!(m.resolve_text(1, 1, c).unwrap(), "1.2");
        let d = m.advance(1, 0).unwrap();
        assert_eq!(m.resolve_text(1, 0, d).unwrap(), "2.");
        let e = m.advance(1, 1).unwrap();
        assert_eq!(m.resolve_text(1, 1, e).unwrap(), "2.1"); // deeper level reset on parent advance
    }

    /// Each level's `%N` is formatted with its OWN numFmt (§17.9.11): an
    /// upper-letter parent with a decimal child renders "A.1".
    #[test]
    fn multilevel_per_level_format() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="2">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1"/></w:lvl>
                 <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>"#);
        m.advance(2, 0).unwrap();
        let c = m.advance(2, 1).unwrap();
        assert_eq!(m.resolve_text(2, 1, c).unwrap(), "A.1");
    }

    /// ECMA-376 §17.9.26 demonstrates that two numIds referencing one abstractNum
    /// share a restart: the sequence 5, 5, 6, 5 produces 1, 2, 1, 2 when numId 6
    /// has startOverride=1. This test exercises the same specified transition.
    #[test]
    fn shared_abstract_counter_restarts_on_start_override() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="20">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="6"><w:abstractNumId w:val="20"/></w:num>
               <w:num w:numId="30"><w:abstractNumId w:val="20"/>
                 <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
               </w:num>"#);
        // Article body (numId=6): 1, 2, 3, 4.
        for expected in ["1.", "2.", "3.", "4."] {
            let c = m.advance(6, 0).unwrap();
            assert_eq!(m.resolve_text(6, 0, c).unwrap(), expected);
        }
        // Masthead heading restarts the shared abstract counter to 1 (numId=30).
        let c = m.advance(30, 0).unwrap();
        assert_eq!(m.resolve_text(30, 0, c).unwrap(), "1.");
        // Body resumes with numId=6 — continues the restarted count: 2, 3, 4.
        for expected in ["2.", "3.", "4."] {
            let c = m.advance(6, 0).unwrap();
            assert_eq!(m.resolve_text(6, 0, c).unwrap(), expected);
        }
    }

    /// A bare `<w:num>` with no override but sharing an abstract with another
    /// num continues the shared count (no accidental per-numId restart).
    #[test]
    fn shared_abstract_counter_continues_across_numids() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="7">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num>
               <w:num w:numId="2"><w:abstractNumId w:val="7"/></w:num>"#);
        let a = m.advance(1, 0).unwrap();
        assert_eq!(m.resolve_text(1, 0, a).unwrap(), "1.");
        let b = m.advance(2, 0).unwrap(); // different numId, same abstract ⇒ continues
        assert_eq!(m.resolve_text(2, 0, b).unwrap(), "2.");
        let c = m.advance(1, 0).unwrap();
        assert_eq!(m.resolve_text(1, 0, c).unwrap(), "3.");
    }

    /// A dangling numId (no `<w:num>`) whose value equals a live abstractNumId
    /// must NOT hijack that abstract's running counter — numId and abstractNumId
    /// are independent ID spaces (§17.9.2 / §17.9.5). Here abstractNumId 4 is
    /// referenced by numId 5; a paragraph then references the unmapped numId 4.
    #[test]
    fn orphan_numid_equal_to_abstract_id_keeps_disjoint_counter() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="4">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="5"><w:abstractNumId w:val="4"/></w:num>"#);
        let a = m.advance(5, 0).unwrap();
        assert_eq!(m.resolve_text(5, 0, a).unwrap(), "1.");
        // numId 4 has no <w:num>; it must start its own count at 1, not read
        // abstractNumId 4's counter (which would yield 2).
        let b = m.advance(4, 0).unwrap();
        assert_eq!(m.resolve_text(4, 0, b).unwrap(), "1.");
    }

    /// ECMA-376 §17.18.59 — the Rust `format_counter` MUST match the core TS
    /// `formatOrdinalNumber` byte-for-byte (list markers resolve here at parse
    /// time). These expected values are the SAME rows as
    /// packages/core/src/text/number-format.test.ts — keep the two in sync.
    #[test]
    fn format_counter_international_matches_core_ts() {
        let cases: &[(&str, &[(u32, &str)])] = &[
            // Latin repeat-letter (beyond 26 now repeats — previously Rust only
            // handled 1–26, a bug this parity pass fixes).
            (
                "upperLetter",
                &[(1, "A"), (26, "Z"), (27, "AA"), (54, "BBB")],
            ),
            (
                "lowerLetter",
                &[(1, "a"), (26, "z"), (27, "aa"), (53, "aaa")],
            ),
            (
                "upperRoman",
                &[(4, "IV"), (123, "CXXIII"), (3999, "MMMCMXCIX")],
            ),
            ("lowerRoman", &[(1, "i"), (123, "cxxiii")]),
            // Positional digit substitution.
            ("decimalFullWidth", &[(1, "１"), (123, "１２３")]),
            ("thaiNumbers", &[(1, "๑"), (123, "๑๒๓")]),
            ("hindiNumbers", &[(1, "१"), (123, "१२३")]),
            (
                "ideographDigital",
                &[(10, "一〇"), (100, "一〇〇"), (2024, "二〇二四")],
            ),
            ("koreanDigital", &[(10, "일영"), (100, "일영영")]),
            // 十-prefix positional.
            (
                "chineseCounting",
                &[
                    (10, "十"),
                    (20, "二十"),
                    (99, "九十九"),
                    (100, "一〇〇"),
                    (101, "一〇一"),
                ],
            ),
            // Grouped counting / legal.
            (
                "japaneseCounting",
                &[
                    (10, "十"),
                    (11, "十一"),
                    (100, "百"),
                    (111, "百十一"),
                    (2024, "二千二十四"),
                    (10005, "一万五"),
                    (12345, "一万二千三百四十五"),
                ],
            ),
            (
                "chineseCountingThousand",
                &[
                    (10, "一十"),
                    (100, "一百"),
                    (1001, "一千零一"),
                    (2024, "二千零二十四"),
                    (10005, "一万零五"),
                ],
            ),
            (
                "chineseLegalSimplified",
                &[
                    (1, "壹"),
                    (10, "壹拾"),
                    (123, "壹佰贰拾叁"),
                    (1001, "壹仟零壹"),
                ],
            ),
            (
                "ideographLegalTraditional",
                &[(10, "壹拾"), (123, "壹佰貳拾參")],
            ),
            (
                "japaneseLegal",
                &[(10, "壱拾"), (2024, "弐阡弐拾四"), (10000, "壱萬")],
            ),
            (
                "koreanCounting",
                &[(10, "십"), (11, "십일"), (2024, "이천이십사")],
            ),
            (
                "koreanLegal",
                &[
                    (1, "하나"),
                    (10, "열"),
                    (11, "열하나"),
                    (21, "스물하나"),
                    (99, "아흔아홉"),
                    (100, "100"),
                ],
            ),
            // Repeat-letter non-Latin alphabets.
            ("arabicAlpha", &[(1, "أ"), (12, "س"), (28, "ي"), (29, "أأ")]),
            ("arabicAbjad", &[(1, "أ"), (12, "ل"), (28, "ظ"), (29, "أأ")]),
            ("russianLower", &[(1, "а"), (29, "я"), (30, "аа")]),
            ("thaiLetters", &[(1, "ก"), (41, "ฮ"), (42, "กก")]),
            ("chosung", &[(1, "ㄱ"), (14, "ㅎ"), (15, "ㄱㄱ")]),
            ("ganada", &[(1, "가"), (14, "하"), (15, "가가")]),
            ("hindiVowels", &[(1, "क"), (37, "ह"), (38, "कक")]),
            (
                "hindiConsonants",
                &[(1, "अ"), (16, "औ"), (17, "अं"), (18, "अः")],
            ),
            // Hebrew: positional gematria / alphabet-with-ת-suffix.
            (
                "hebrew1",
                &[
                    (1, "א"),
                    (15, "טו"),
                    (16, "טז"),
                    (17, "יז"),
                    (123, "קכג"),
                    (500, "ך"),
                ],
            ),
            // hebrew2: result glyph once + ת per 22 subtracted (§17.18.59 steps;
            // §17.16.4.3.1 example 123 → מ + 5×ת). NOT the repeat scheme.
            (
                "hebrew2",
                &[
                    (1, "א"),
                    (22, "ת"),
                    (23, "את"),
                    (24, "בת"),
                    (44, "תת"),
                    (45, "אתת"),
                    (123, "מתתתתת"),
                ],
            ),
            // Other algorithmic systems.
            (
                "hex",
                &[(1, "1"), (10, "A"), (16, "10"), (31, "1F"), (255, "FF")],
            ),
            (
                "numberInDash",
                &[(1, "- 1 -"), (10, "- 10 -"), (123, "- 123 -")],
            ),
            (
                "decimalZero",
                &[(1, "01"), (9, "09"), (10, "10"), (100, "100")],
            ),
            // Enclosed decimals + katakana a-i-u-e-o sequences (§17.18.59).
            // decimalEnclosedCircle: 1–20 tabled (①..⑳); 21+ falls back to
            // decimal per the spec example "…, ⑲, ⑳, 21, …". Boundary 20/21.
            (
                "decimalEnclosedCircle",
                &[(1, "①"), (20, "⑳"), (21, "21"), (100, "100")],
            ),
            // aiueoFullWidth: 48-entry enumerated set incl. archaic ヰ/ヱ, so wo/n
            // sit at 47/48; repeat past 48.
            (
                "aiueoFullWidth",
                &[
                    (1, "ア"),
                    (44, "ワ"),
                    (45, "ヰ"),
                    (46, "ヱ"),
                    (47, "ヲ"),
                    (48, "ン"),
                    (49, "アア"),
                    (97, "アアア"),
                ],
            ),
            // aiueo: half-width katakana, 46-entry set (no archaic forms), wo/n at
            // 45/46; repeat past 46.
            (
                "aiueo",
                &[(1, "ｱ"), (44, "ﾜ"), (45, "ｦ"), (46, "ﾝ"), (47, "ｱｱ")],
            ),
            // Documented residual / spell-outs fall back to decimal.
            ("cardinalText", &[(5, "5")]),
            ("thaiCounting", &[(5, "5")]),
            ("none", &[(0, ""), (1, ""), (5, "")]),
        ];
        for (fmt, rows) in cases {
            for (input, expected) in *rows {
                assert_eq!(
                    format_counter(*input, fmt),
                    *expected,
                    "format_counter({input}, {fmt:?})"
                );
            }
        }
    }

    /// A numbered list whose level uses an international `<w:numFmt>` resolves the
    /// marker through `format_counter` — the end-to-end list path (§17.9.17).
    #[test]
    fn list_marker_uses_international_numfmt() {
        let mut m = map(r#"<w:abstractNum w:abstractNumId="9">
                 <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="ideographDigital"/><w:lvlText w:val="%1."/></w:lvl>
               </w:abstractNum>
               <w:num w:numId="9"><w:abstractNumId w:val="9"/></w:num>"#);
        let a = m.advance(9, 0).unwrap();
        assert_eq!(m.resolve_text(9, 0, a).unwrap(), "一."); // 1 → 一
        let b = m.advance(9, 0).unwrap();
        assert_eq!(m.resolve_text(9, 0, b).unwrap(), "二."); // 2 → 二
    }

    #[test]
    fn clones_share_first_failure_but_keep_independent_counter_state() {
        let mut original = map(
            r#"<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>"#,
        );
        let mut cloned = original.clone();
        assert_eq!(original.advance(1, 0), Ok(1));
        assert_eq!(cloned.advance(1, 0), Ok(1));
        original.record_counter_error(CounterError::Overflow);
        cloned.record_counter_error(CounterError::InvalidLevel);
        assert_eq!(
            original.check_counter_error(),
            Err("numbering counter overflow".to_string())
        );
        assert_eq!(cloned.check_counter_error(), original.check_counter_error());
    }
}
