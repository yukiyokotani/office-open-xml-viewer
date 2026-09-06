use crate::styles::{parse_run_fmt, RunFmt};
use crate::xml_util::*;
use ooxml_common::blip::mime_from_ext;
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::{attr_ns, relationships};
use std::collections::{HashMap, HashSet};

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

/// Key into the running-counter map. numId values and abstractNumId values are
/// independent ID sequences in WordprocessingML (§17.9.2 / §17.9.5), so they
/// must NOT share a `u32` key space: a dangling numId (no `<w:num>`) that
/// happens to equal a real abstractNumId would otherwise hijack that abstract's
/// live count. `Abstract` holds the shared count for a resolved num; `OrphanNum`
/// gives an unresolved num its own disjoint counter.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum CounterKey {
    Abstract(u32),
    OrphanNum(u32),
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
    /// evidence (see `counter_instance_tests`). Keyed by `CounterKey` so an
    /// unresolved numId gets a disjoint counter instead of colliding with an
    /// abstractNumId.
    counters: HashMap<CounterKey, HashMap<u32, u32>>,
    /// (numId, level) pairs already advanced at least once. A numId carrying a
    /// `<w:lvlOverride><w:startOverride>` restarts the shared abstract counter
    /// only on its FIRST appearance at that level. ECMA-376 §17.9.26 defines
    /// `startOverride` and demonstrates its reset propagating across numIds on
    /// one abstractNum; applying it only once per numId follows measured
    /// native-Word compatibility behavior.
    started: HashSet<(u32, u32)>,
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
    /// lists at 9 levels (ST_Ilvl, §17.18.38).
    pub fn level_for_style(&self, num_id: u32, style_id: &str) -> Option<u32> {
        (0..9).find(|&l| {
            self.get_level(num_id, l)
                .is_some_and(|def| def.p_style.as_deref() == Some(style_id))
        })
    }

    pub fn get_start(&self, num_id: u32, level: u32) -> u32 {
        if let Some(ov) = self.num_overrides.get(&num_id).and_then(|m| m.get(&level)) {
            return *ov;
        }
        self.get_level(num_id, level).map(|l| l.start).unwrap_or(1)
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
    pub fn advance(&mut self, num_id: u32, level: u32) -> u32 {
        // Pre-compute start values to avoid borrow conflicts. `get_start`
        // already folds in any per-numId `<w:startOverride>` for the level.
        let starts: Vec<u32> = (0..=level).map(|l| self.get_start(num_id, l)).collect();
        let key = self.counter_key(num_id);
        let has_override = self
            .num_overrides
            .get(&num_id)
            .is_some_and(|m| m.contains_key(&level));
        // `insert` returns true when the pair was NOT already present.
        let first_for_num = self.started.insert((num_id, level));

        // Resolve each live descendant's own policy, including a complete
        // level replacement. A never-restarting parent does not shield its
        // children, and merely seeding an ancestor is not an occurrence of it.
        // Iterate existing counters, never an input-provided restart range.
        let resets: Vec<u32> = self
            .counters
            .get(&key)
            .into_iter()
            .flat_map(|counts| counts.keys().copied())
            .filter(|&deeper| {
                let threshold = self
                    .get_level(num_id, deeper)
                    .and_then(|def| def.restart)
                    .filter(|&value| value <= deeper)
                    .unwrap_or(deeper);
                deeper > level && level < threshold
            })
            .collect();

        let entry = self.counters.entry(key).or_default();

        for k in resets {
            entry.remove(&k);
        }

        // Seed shallower levels to their start (their displayed value when they
        // are never advanced themselves) — but never clobber a live ancestor.
        for (lvl, &start) in starts.iter().enumerate().take(level as usize) {
            entry.entry(lvl as u32).or_insert(start);
        }

        // A startOverride restarts the shared counter on first use of this num;
        // otherwise the level shows `start` on its first appearance on the
        // abstract and increments thereafter.
        let val = if first_for_num && has_override {
            starts[level as usize]
        } else {
            match entry.get(&level) {
                Some(&v) => v + 1,
                None => starts[level as usize],
            }
        };
        entry.insert(level, val);
        val
    }

    /// The counter-map key for a numId: the shared `Abstract(abstractNumId)`
    /// when the `<w:num>` resolves, else an `OrphanNum(numId)` in a disjoint key
    /// space so a dangling numId can never collide with a real abstractNumId's
    /// running counter.
    fn counter_key(&self, num_id: u32) -> CounterKey {
        match self.num_to_abstract.get(&num_id) {
            Some(&abs) => CounterKey::Abstract(abs),
            None => CounterKey::OrphanNum(num_id),
        }
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
    pub fn resolve_text(&self, num_id: u32, level: u32, counter: u32) -> String {
        let Some(lvl) = self.get_level(num_id, level) else {
            return format!("{}.", counter);
        };
        let key = self.counter_key(num_id);

        let mut text = lvl.text.clone();
        // Replace from the deepest placeholder down so "%1" can never partially
        // match a two-digit "%1N" (Word caps lists at 9 levels, so this is
        // belt-and-braces — but cheap).
        for k in (0..=level).rev() {
            let val = if k == level {
                counter
            } else {
                self.counters
                    .get(&key)
                    .and_then(|m| m.get(&k))
                    .copied()
                    .unwrap_or_else(|| self.get_start(num_id, k))
            };
            // 17.9.4 applies to this marker's entire displayed level text,
            // including its own placeholder. Keep authored formats intact so
            // other markers continue to use their own definitions. MS-OE376
            // 2.1.280(b) documents Word retaining `none`; this path follows the
            // normative decimal rule, without a format-specific exception.
            let fmt = if lvl.legal {
                "decimal"
            } else {
                self.get_level(num_id, k)
                    .map(|l| l.format.as_str())
                    .unwrap_or(lvl.format.as_str())
            };
            text = text.replace(&format!("%{}", k + 1), &format_counter(val, fmt));
        }
        text
    }
}

// ── ST_NumberFormat rendering (ECMA-376 §17.18.59) ──────────────────────────
// This is the RUST twin of `packages/core/src/text/number-format.ts`
// (`formatOrdinalNumber`). List markers resolve to a final string at PARSE time
// (`resolve_text` composes `%1.%2` here in Rust), so the same numbering systems
// must be implemented on both sides and produce BYTE-IDENTICAL output. When you
// touch a format here, mirror it there (and vice versa); the TS unit tests are
// the reference values. `bullet` is a list-only concern (no §17.18.59 numeric
// meaning) and stays Rust-only.
fn format_counter(n: u32, format: &str) -> String {
    // ECMA-376 17.18.59: `none` suppresses the number, including start=0.
    // Mirror the shared TS field formatter instead of using decimal fallback.
    if format == "none" {
        return String::new();
    }
    if format == "bullet" {
        return "•".to_string();
    }
    // The numeric systems are 1-based; a level with start=0 (rare) or an
    // underflow falls back to the decimal string, matching the TS `n >= 1` gate.
    if n == 0 {
        return n.to_string();
    }
    match format {
        "decimal" | "decimalHalfWidth" => n.to_string(),
        // Roman.
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        // Latin + non-Latin repeat-letter alphabets (§17.18.59).
        "lowerLetter" => repeat_alphabet(n, &latin_upper()).to_lowercase(),
        "upperLetter" => repeat_alphabet(n, &latin_upper()),
        "arabicAlpha" => repeat_alphabet(n, ARABIC_ALPHA),
        "arabicAbjad" => repeat_alphabet(n, ARABIC_ABJAD),
        "russianLower" => repeat_alphabet(n, RUSSIAN_LOWER),
        "russianUpper" => repeat_alphabet(n, RUSSIAN_UPPER),
        "thaiLetters" => repeat_alphabet(n, THAI_LETTERS),
        "chosung" => repeat_alphabet(n, KOREAN_CHOSUNG),
        "ganada" => repeat_alphabet(n, KOREAN_GANADA),
        "hindiVowels" => repeat_alphabet(n, HINDI_VOWELS),
        "hindiConsonants" => repeat_alphabet(n, HINDI_CONSONANTS),
        // Katakana a-i-u-e-o sequences (repeat scheme, like the letter alphabets).
        "aiueoFullWidth" => repeat_alphabet(n, KATAKANA_FULLWIDTH),
        "aiueo" => repeat_alphabet(n, KATAKANA_HALFWIDTH),
        // Enclosed decimals (bounded set → decimal fallback past the range).
        "decimalEnclosedCircle" => to_enclosed_circle(n),
        // Hebrew: positional gematria / alphabet-with-ת-suffix (NOT repeat).
        "hebrew1" => to_hebrew_gematria(n),
        "hebrew2" => to_hebrew2(n),
        // Other algorithmic systems (§17.18.59 hex / numberInDash / decimalZero).
        "hex" => format!("{:X}", n),
        "numberInDash" => format!("- {} -", n),
        "decimalZero" => {
            if n <= 9 {
                format!("0{}", n)
            } else {
                n.to_string()
            }
        }
        // Positional digit substitution.
        "decimalFullWidth" => to_positional_digits(n, DIGITS_FULLWIDTH),
        "thaiNumbers" => to_positional_digits(n, DIGITS_THAI),
        "hindiNumbers" => to_positional_digits(n, DIGITS_HINDI),
        "ideographDigital" | "japaneseDigitalTenThousand" => {
            to_positional_digits(n, DIGITS_IDEOGRAPH)
        }
        "koreanDigital" => to_positional_digits(n, DIGITS_KOREAN),
        "koreanDigital2" => to_positional_digits(n, DIGITS_KOREAN2),
        "taiwaneseDigital" => to_positional_digits(n, DIGITS_TAIWANESE),
        // 十-prefix positional.
        "chineseCounting" => to_chinese_counting(n, DIGITS_IDEOGRAPH),
        "taiwaneseCounting" => to_chinese_counting(n, DIGITS_TAIWANESE),
        // Grouped counting / legal CJK.
        "japaneseCounting" => to_myriad_grouped(n, &MYRIAD_JAPANESE),
        "chineseCountingThousand" => to_myriad_grouped(n, &MYRIAD_CHINESE),
        "taiwaneseCountingThousand" => to_myriad_grouped(n, &MYRIAD_CHINESE),
        "chineseLegalSimplified" => to_myriad_grouped(n, &MYRIAD_CHINESE_LEGAL),
        "ideographLegalTraditional" => to_myriad_grouped(n, &MYRIAD_TRAD_LEGAL),
        "japaneseLegal" => to_myriad_grouped(n, &MYRIAD_JAPANESE_LEGAL),
        "koreanCounting" => to_myriad_grouped(n, &MYRIAD_KOREAN),
        "koreanLegal" => to_korean_legal(n),
        // Documented residual (language spell-outs / unimplemented) → decimal.
        _ => n.to_string(),
    }
}

// §17.18.59 koreanLegal — native-Korean tens-word + ones-word, tabled for 1–99
// (≥100 undefined by the spec → decimal fallback). Mirrors TS `toKoreanLegal`.
const KOREAN_LEGAL_ONES: &[&str] = &[
    "", "하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉",
];
const KOREAN_LEGAL_TENS: &[&str] = &[
    "", "열", "스물", "서른", "마흔", "쉰", "예순", "일흔", "여든", "아흔",
];

fn to_korean_legal(n: u32) -> String {
    if n >= 100 {
        return n.to_string();
    }
    let tens = n / 10;
    let ones = n % 10;
    format!(
        "{}{}",
        KOREAN_LEGAL_TENS[tens as usize], KOREAN_LEGAL_ONES[ones as usize]
    )
}

fn to_roman(n: u32) -> String {
    let vals = [
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
    let mut n = n;
    let mut s = String::new();
    for (v, r) in &vals {
        while n >= *v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

// A–Z, built for the Latin letter converters (repeat scheme, not base-26).
fn latin_upper() -> Vec<&'static str> {
    const A: [&str; 26] = [
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R",
        "S", "T", "U", "V", "W", "X", "Y", "Z",
    ];
    A.to_vec()
}

// §17.18.59 arabicAlpha "Arabic Alphabet" — positions 1–28.
const ARABIC_ALPHA: &[&str] = &[
    "أ", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر", "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ",
    "ف", "ق", "ك", "ل", "م", "ن", "ه", "و", "ي",
];
// §17.18.59 arabicAbjad "Arabic Abjad Numerals" — positions 1–28.
const ARABIC_ABJAD: &[&str] = &[
    "أ", "ب", "ج", "د", "ه", "و", "ز", "ح", "ط", "ي", "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق",
    "ر", "ش", "ت", "ث", "خ", "ذ", "ض", "غ", "ظ",
];
// §17.18.59 hebrew2 "Hebrew Alphabet" — positions 1–22.
const HEBREW_ALPHABET: &[&str] = &[
    "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ", "ק",
    "ר", "ש", "ת",
];
// §17.18.59 russianLower/Upper — positions 1–29 (alphabet minus ё, й, ъ, ь).
const RUSSIAN_LOWER: &[&str] = &[
    "а", "б", "в", "г", "д", "е", "ж", "з", "и", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у",
    "ф", "х", "ц", "ч", "ш", "щ", "ы", "э", "ю", "я",
];
const RUSSIAN_UPPER: &[&str] = &[
    "А", "Б", "В", "Г", "Д", "Е", "Ж", "З", "И", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У",
    "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ы", "Э", "Ю", "Я",
];
// §17.18.59 thaiLetters — positions 1–41 (U+0E01, U+0E02, U+0E04, U+0E07–U+0E23,
// U+0E25, U+0E27–U+0E2E).
const THAI_LETTERS: &[&str] = &[
    "ก", "ข", "ค", "ง", "จ", "ฉ", "ช", "ซ", "ฌ", "ญ", "ฎ", "ฏ", "ฐ", "ฑ", "ฒ", "ณ", "ด", "ต", "ถ",
    "ท", "ธ", "น", "บ", "ป", "ผ", "ฝ", "พ", "ฟ", "ภ", "ม", "ย", "ร", "ล", "ว", "ศ", "ษ", "ส", "ห",
    "ฬ", "อ", "ฮ",
];
// §17.18.59 chosung "Korean Chosung" — positions 1–14.
const KOREAN_CHOSUNG: &[&str] = &[
    "ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ",
];
// §17.18.59 ganada "Korean Ganada" — positions 1–14.
const KOREAN_GANADA: &[&str] = &[
    "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하",
];
// §17.18.59 hindiVowels — positions 1–37 = U+0915–U+0939 (contiguous).
const HINDI_VOWELS: &[&str] = &[
    "क", "ख", "ग", "घ", "ङ", "च", "छ", "ज", "झ", "ञ", "ट", "ठ", "ड", "ढ", "ण", "त", "थ", "द", "ध",
    "न", "ऩ", "प", "फ", "ब", "भ", "म", "य", "र", "ऱ", "ल", "ळ", "ऴ", "व", "श", "ष", "स", "ह",
];
// §17.18.59 hindiConsonants — positions 1–18 = U+0905–U+0914 then अं / अः.
const HINDI_CONSONANTS: &[&str] = &[
    "अ", "आ", "इ", "ई", "उ", "ऊ", "ऋ", "ऌ", "ऍ", "ऎ", "ए", "ऐ", "ऑ", "ऒ", "ओ", "औ", "अं", "अः",
];

/// §17.18.59 repeat-letter alphabets: map 1..N into the set; for n>N repeat the
/// SAME glyph once per full N subtracted (not base-N). Mirrors the TS
/// `repeatAlphabet`.
fn repeat_alphabet(n: u32, glyphs: &[&str]) -> String {
    let size = glyphs.len() as u32;
    let repeats = (n - 1) / size + 1;
    let glyph = glyphs[((n - 1) % size) as usize];
    glyph.repeat(repeats as usize)
}

// §17.18.59 aiueoFullWidth "AIUEO Order Full-Width Katakana" — the full-width
// katakana in a-i-u-e-o order, using the SAME repeat scheme as the letter
// alphabets. §17.18.59 ENUMERATES these 48 code points (incl. the archaic ヰ
// U+30F0 and ヱ U+30F1, so wo/n land at 47/48). The section's "positions 1–46"
// prose is a copy/paste artifact from the half-width `aiueo` entry; we follow the
// explicit enumerated character list. Katakana (not hiragana) is also what Word
// emits: [MS-OE376] §2.1.580 note (b) records that where the (1st-edition Part 4)
// standard said "hiragana characters", Word uses katakana — matching the
// 5th-edition code-point list implemented here. Mirrors TS `KATAKANA_FULLWIDTH`.
const KATAKANA_FULLWIDTH: &[&str] = &[
    "\u{30A2}", "\u{30A4}", "\u{30A6}", "\u{30A8}", "\u{30AA}", "\u{30AB}", "\u{30AD}", "\u{30AF}",
    "\u{30B1}", "\u{30B3}", "\u{30B5}", "\u{30B7}", "\u{30B9}", "\u{30BB}", "\u{30BD}", "\u{30BF}",
    "\u{30C1}", "\u{30C4}", "\u{30C6}", "\u{30C8}", "\u{30CA}", "\u{30CB}", "\u{30CC}", "\u{30CD}",
    "\u{30CE}", "\u{30CF}", "\u{30D2}", "\u{30D5}", "\u{30D8}", "\u{30DB}", "\u{30DE}", "\u{30DF}",
    "\u{30E0}", "\u{30E1}", "\u{30E2}", "\u{30E4}", "\u{30E6}", "\u{30E8}", "\u{30E9}", "\u{30EA}",
    "\u{30EB}", "\u{30EC}", "\u{30ED}", "\u{30EF}", "\u{30F0}", "\u{30F1}", "\u{30F2}", "\u{30F3}",
];

// §17.18.59 aiueo "AIUEO Order Half-Width Katakana" — positions 1–46 =
// U+FF71–U+FF9C (ｱ..ﾜ), then U+FF66 (ｦ), then U+FF9D (ﾝ). No archaic ヰ/ヱ (no
// half-width forms). Katakana per [MS-OE376] §2.1.580 note (b) — see
// `KATAKANA_FULLWIDTH` above. Mirrors TS `KATAKANA_HALFWIDTH`.
const KATAKANA_HALFWIDTH: &[&str] = &[
    "\u{FF71}", "\u{FF72}", "\u{FF73}", "\u{FF74}", "\u{FF75}", "\u{FF76}", "\u{FF77}", "\u{FF78}",
    "\u{FF79}", "\u{FF7A}", "\u{FF7B}", "\u{FF7C}", "\u{FF7D}", "\u{FF7E}", "\u{FF7F}", "\u{FF80}",
    "\u{FF81}", "\u{FF82}", "\u{FF83}", "\u{FF84}", "\u{FF85}", "\u{FF86}", "\u{FF87}", "\u{FF88}",
    "\u{FF89}", "\u{FF8A}", "\u{FF8B}", "\u{FF8C}", "\u{FF8D}", "\u{FF8E}", "\u{FF8F}", "\u{FF90}",
    "\u{FF91}", "\u{FF92}", "\u{FF93}", "\u{FF94}", "\u{FF95}", "\u{FF96}", "\u{FF97}", "\u{FF98}",
    "\u{FF99}", "\u{FF9A}", "\u{FF9B}", "\u{FF9C}", "\u{FF66}", "\u{FF9D}",
];

/// §17.18.59 decimalEnclosedCircle: the spec tables 1–20 → U+2460–U+2473 (①..⑳)
/// and states that "for values greater than the size of the set, the items fall
/// back to the decimal format" (its example: …, ⑲, ⑳, 21, …). Unicode does carry
/// follow-on enclosed-number blocks (㉑..㉟ U+3251–U+325F, ㊱..㊿ U+32B1–U+32BF),
/// but neither §17.18.59 nor the Word implementation notes ([MS-OE376] §2.1.580,
/// which records this section's sibling deviations in detail) documents Word
/// continuing the circled sequence past 20 — so we stay with the specified
/// decimal fallback at 21+ until primary evidence (real Word output or an
/// implementation note) shows otherwise. Mirrors TS `toEnclosedCircle`. Caller
/// guarantees n ≥ 1 (n = 0 is handled by the early return in `format_counter`).
fn to_enclosed_circle(n: u32) -> String {
    match n {
        1..=20 => char::from_u32(0x2460 + (n - 1))
            .map(String::from)
            .unwrap_or_else(|| n.to_string()),
        _ => n.to_string(), // 21+ : §17.18.59 decimal fallback.
    }
}

// Positional digit sets, index 0 = zero glyph … index 9 (§17.18.59).
const DIGITS_FULLWIDTH: &[&str] = &["０", "１", "２", "３", "４", "５", "６", "７", "８", "９"];
const DIGITS_THAI: &[&str] = &["๐", "๑", "๒", "๓", "๔", "๕", "๖", "๗", "๘", "๙"];
const DIGITS_HINDI: &[&str] = &["०", "१", "२", "३", "४", "५", "६", "७", "८", "९"];
const DIGITS_IDEOGRAPH: &[&str] = &["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const DIGITS_KOREAN: &[&str] = &["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
const DIGITS_KOREAN2: &[&str] = &["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const DIGITS_TAIWANESE: &[&str] = &["○", "一", "二", "三", "四", "五", "六", "七", "八", "九"];

/// §17.18.59 base-10 positional digit substitution. Mirrors TS `toPositionalDigits`.
fn to_positional_digits(n: u32, digits: &[&str]) -> String {
    n.to_string()
        .bytes()
        .map(|b| digits[(b - b'0') as usize])
        .collect()
}

/// §17.18.59 chineseCounting / taiwaneseCounting: base-10 positional with the 十
/// tens-word for 2-digit values only (10 → 十, 20 → 二十, 99 → 九十九), pure
/// positional at ≥ 100 (100 → 一〇〇). Mirrors TS `toChineseCounting`.
fn to_chinese_counting(n: u32, digits: &[&str]) -> String {
    if n < 10 {
        return digits[n as usize].to_string();
    }
    if n < 100 {
        let tens = n / 10;
        let ones = n % 10;
        let head = if tens == 1 {
            "十".to_string()
        } else {
            format!("{}十", digits[tens as usize])
        };
        return if ones == 0 {
            head
        } else {
            format!("{}{}", head, digits[ones as usize])
        };
    }
    to_positional_digits(n, digits)
}

// ── Grouped CJK counting / legal (myriad grouping) ──────────────────────────
// Mirrors the TS `MyriadTable` + `toMyriadGrouped`.
struct MyriadTable {
    digits: &'static [&'static str], // index 0 = 零/〇 zero-fill glyph … 9
    ten: &'static str,
    hundred: &'static str,
    thousand: &'static str,
    myriad: &'static str,
    elide_one: bool,   // Japanese/Korean elide the "1" before 十/百/千.
    insert_zero: bool, // Chinese counting/legal fill an interior gap with 零.
}

const CJK_DIGITS: &[&str] = DIGITS_KOREAN2; // 零 一 二 … 九
const MYRIAD_JAPANESE: MyriadTable = MyriadTable {
    digits: CJK_DIGITS,
    ten: "十",
    hundred: "百",
    thousand: "千",
    myriad: "万",
    elide_one: true,
    insert_zero: false,
};
const MYRIAD_CHINESE: MyriadTable = MyriadTable {
    elide_one: false,
    insert_zero: true,
    ..MYRIAD_JAPANESE
};
const MYRIAD_KOREAN: MyriadTable = MyriadTable {
    digits: &["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"],
    ten: "십",
    hundred: "백",
    thousand: "천",
    myriad: "만",
    elide_one: true,
    insert_zero: false,
};
const MYRIAD_CHINESE_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"],
    ten: "拾",
    hundred: "佰",
    thousand: "仟",
    myriad: "万",
    elide_one: false,
    insert_zero: true,
};
const MYRIAD_JAPANESE_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壱", "弐", "参", "四", "伍", "六", "七", "八", "九"],
    ten: "拾",
    hundred: "百",
    thousand: "阡",
    myriad: "萬",
    elide_one: false,
    insert_zero: false,
};
const MYRIAD_TRAD_LEGAL: MyriadTable = MyriadTable {
    digits: &["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"],
    ten: "拾",
    hundred: "佰",
    thousand: "仟",
    myriad: "萬",
    elide_one: false,
    insert_zero: false,
};

/// Render one 4-digit myriad group (0–9999). Mirrors TS `renderMyriadGroup`.
fn render_myriad_group(group: u32, t: &MyriadTable) -> String {
    let thousands = group / 1000 % 10;
    let hundreds = group / 100 % 10;
    let tens = group / 10 % 10;
    let ones = group % 10;
    let places = [
        (thousands, t.thousand),
        (hundreds, t.hundred),
        (tens, t.ten),
        (ones, ""),
    ];
    let mut out = String::new();
    let mut saw_non_zero = false;
    let mut pending_zero = false;
    for (digit, unit) in places {
        if digit == 0 {
            if saw_non_zero {
                pending_zero = true;
            }
            continue;
        }
        if pending_zero {
            if t.insert_zero {
                out.push_str(t.digits[0]);
            }
            pending_zero = false;
        }
        if t.elide_one && digit == 1 && !unit.is_empty() {
            out.push_str(unit);
        } else {
            out.push_str(t.digits[digit as usize]);
            out.push_str(unit);
        }
        saw_non_zero = true;
    }
    out
}

/// East-Asian myriad-grouped counting/legal formatter. Mirrors TS
/// `toMyriadGrouped` (values ≥ 10^8 recurse through 億).
fn to_myriad_grouped(n: u32, t: &MyriadTable) -> String {
    if n >= 100_000_000 {
        let upper = n / 100_000_000;
        let lower = n % 100_000_000;
        let head = format!("{}億", to_myriad_grouped(upper, t));
        if lower == 0 {
            return head;
        }
        let gap = if t.insert_zero && lower < 10_000_000 {
            t.digits[0]
        } else {
            ""
        };
        return format!("{}{}{}", head, gap, to_myriad_grouped(lower, t));
    }
    let upper_group = n / 10000;
    let lower_group = n % 10000;
    let mut out = String::new();
    if upper_group > 0 {
        out.push_str(&render_myriad_group(upper_group, t));
        out.push_str(t.myriad);
    }
    if lower_group > 0 {
        if t.insert_zero && upper_group > 0 && lower_group < 1000 {
            out.push_str(t.digits[0]);
        }
        out.push_str(&render_myriad_group(lower_group, t));
    }
    out
}

// §17.18.59 hebrew1 gematria. Mirrors TS `toHebrewGematria`.
const HEBREW_ONES: &[&str] = &["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
const HEBREW_TENS: &[&str] = &["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
const HEBREW_HUNDREDS: &[&str] = &["", "ק", "ר", "ש", "ת", "ך", "ם", "ן", "ף", "ץ"];

fn to_hebrew_gematria(n: u32) -> String {
    let mut out = String::new();
    let mut rem = n;
    let thousands = rem / 1000;
    rem %= 1000;
    let hundreds = rem / 100;
    rem %= 100;
    if thousands > 0 {
        out.push_str(HEBREW_ONES[(thousands % 10) as usize]);
    }
    out.push_str(HEBREW_HUNDREDS[hundreds as usize]);
    if rem == 15 {
        out.push_str("טו");
        return out;
    }
    if rem == 16 {
        out.push_str("טז");
        return out;
    }
    let tens = rem / 10;
    let ones = rem % 10;
    out.push_str(HEBREW_TENS[tens as usize]);
    out.push_str(HEBREW_ONES[ones as usize]);
    out
}

/// §17.18.59 hebrew2 — NOT the repeat-letter scheme: subtract 22 until the
/// result is ≤ 22, write THAT glyph once, then append ת once per subtraction
/// (23 → את, 24 → בת; §17.16.4.3.1 field example 123 → מ + 5×ת). Mirrors TS
/// `toHebrew2`.
fn to_hebrew2(n: u32) -> String {
    let size = HEBREW_ALPHABET.len() as u32; // 22
    let subtractions = (n - 1) / size;
    let remainder = n - size * subtractions; // 1..=22
    format!(
        "{}{}",
        HEBREW_ALPHABET[(remainder - 1) as usize],
        "ת".repeat(subtractions as usize)
    )
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
        let c1 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c1), "3.1");
        let c2 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c2), "3.2");
        let c3 = m.advance(5, 1);
        assert_eq!(m.resolve_text(5, 1, c3), "3.3");
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
        let a = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, a), "1.");
        let b = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, b), "1.1");
        let c = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, c), "1.2");
        let d = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, d), "2.");
        let e = m.advance(1, 1);
        assert_eq!(m.resolve_text(1, 1, e), "2.1"); // deeper level reset on parent advance
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
        m.advance(2, 0);
        let c = m.advance(2, 1);
        assert_eq!(m.resolve_text(2, 1, c), "A.1");
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
            let c = m.advance(6, 0);
            assert_eq!(m.resolve_text(6, 0, c), expected);
        }
        // Masthead heading restarts the shared abstract counter to 1 (numId=30).
        let c = m.advance(30, 0);
        assert_eq!(m.resolve_text(30, 0, c), "1.");
        // Body resumes with numId=6 — continues the restarted count: 2, 3, 4.
        for expected in ["2.", "3.", "4."] {
            let c = m.advance(6, 0);
            assert_eq!(m.resolve_text(6, 0, c), expected);
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
        let a = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, a), "1.");
        let b = m.advance(2, 0); // different numId, same abstract ⇒ continues
        assert_eq!(m.resolve_text(2, 0, b), "2.");
        let c = m.advance(1, 0);
        assert_eq!(m.resolve_text(1, 0, c), "3.");
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
        let a = m.advance(5, 0);
        assert_eq!(m.resolve_text(5, 0, a), "1.");
        // numId 4 has no <w:num>; it must start its own count at 1, not read
        // abstractNumId 4's counter (which would yield 2).
        let b = m.advance(4, 0);
        assert_eq!(m.resolve_text(4, 0, b), "1.");
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
        let a = m.advance(9, 0);
        assert_eq!(m.resolve_text(9, 0, a), "一."); // 1 → 一
        let b = m.advance(9, 0);
        assert_eq!(m.resolve_text(9, 0, b), "二."); // 2 → 二
    }
}
