//! Text body / paragraph / run parsing plus the list-style level-* helpers
//! (`LevelFontSizes` / `LevelIndents` / `LevelBullets` and their read/extract/
//! has/merge functions, shared with the master extractors in `lib.rs`).
//! Extracted verbatim from `lib.rs`. Shared XML helpers (`child`,
//! `children_vec`, `attr`, `attr_r`, `attr_i64`, `attr_f64`, `resolve_path`)
//! stay in `lib.rs`; the colour + theme helpers live in `fill` / `theme`.

use crate::fill::{parse_color_node, parse_fill, parse_reflection, parse_shadow};
use crate::script_font::{resolve_slot_face, theme_token_set, FontSlot};
use crate::theme::resolve_theme_typeface;
use crate::types::*;
use crate::{attr, attr_f64, attr_i64, attr_r, child, children_vec, resolve_path, PptxZip};
use ooxml_common::blip::mime_from_ext;
use ooxml_common::math::parse_omath_nodes;
use ooxml_common::text::{parse_lnspc, SpaceLine};
use ooxml_common::units::text_point_to_pt;
use std::collections::{BTreeMap, HashMap};

type PropertyAttributes = BTreeMap<String, String>;

#[derive(Clone, serde::Serialize)]
struct InheritedRelationship {
    target: String,
    /// Set for master/layout levels. Slide-local targets retain the existing
    /// slide-relative representation used by the viewer.
    source_part: Option<String>,
}

fn merge_attributes(higher: &PropertyAttributes, lower: &PropertyAttributes) -> PropertyAttributes {
    let mut merged = lower.clone();
    merged.extend(higher.clone());
    merged
}

fn xml_attributes(node: roxmltree::Node<'_, '_>) -> PropertyAttributes {
    node.attributes()
        .filter(|a| !matches!(a.name(), "dirty" | "err"))
        .map(|a| (a.name().to_owned(), a.value().to_owned()))
        .collect()
}

/// One `CT_TextSpacing` choice from `<a:spcBef>` / `<a:spcAft>` (ECMA-376
/// §21.1.2.2.9-.10): an absolute `<a:spcPts>` in hundredths of a point or a
/// `<a:spcPct>` in thousandths of a percent of the text size.
#[derive(Clone, Copy, Debug, PartialEq, serde::Serialize)]
pub(crate) enum ParagraphSpacing {
    Points(i64),
    Percent(f64),
}

/// Read `<a:spcBef>` / `<a:spcAft>` under a paragraph-properties node.
pub(crate) fn paragraph_spacing(
    properties: roxmltree::Node<'_, '_>,
    name: &str,
) -> Option<ParagraphSpacing> {
    let spacing = child(properties, name)?;
    child(spacing, "spcPts")
        .and_then(|n| attr_i64(&n, "val"))
        .map(ParagraphSpacing::Points)
        .or_else(|| {
            child(spacing, "spcPct")
                .and_then(|n| attr_f64(&n, "val"))
                .map(ParagraphSpacing::Percent)
        })
}

impl ParagraphSpacing {
    /// Split into the model's exclusive (points, percent) fields.
    pub(crate) fn split(value: Option<Self>) -> (Option<i64>, Option<f64>) {
        match value {
            Some(Self::Points(v)) => (Some(v), None),
            Some(Self::Percent(v)) => (None, Some(v)),
            None => (None, None),
        }
    }
}

/// Per-list-level paragraph spacing from `<a:lvlNpPr>`: `spcBef`, `spcAft`
/// and `lnSpc`, each a percentage or points (CT_TextSpacing). Index 0..=8 → lvl1pPr..
/// lvl9pPr. Each level and property inherits independently (ECMA-376
/// §21.1.2.4): a paragraph at level N takes level N of the nearest list style
/// that sets it, never another level's value. Observed (#1630): a title whose
/// titleStyle set 90 % line spacing on level 1 only laid level-2 paragraphs
/// out at single spacing, and body levels 2-5 took their own 5 pt space before
/// rather than level 1's 10 pt.
#[derive(Clone, Debug, Default, PartialEq, serde::Serialize)]
pub(crate) struct LevelSpacing {
    pub(crate) before: [Option<ParagraphSpacing>; 9],
    pub(crate) after: [Option<ParagraphSpacing>; 9],
    pub(crate) line: [Option<SpaceLine>; 9],
}

impl LevelSpacing {
    /// Read levels 1..9 from a node holding `<a:lvlNpPr>` children: a txBody's
    /// `<a:lstStyle>` or a master `<p:txStyles>` style node.
    pub(crate) fn read(list_style: roxmltree::Node<'_, '_>) -> Self {
        let mut out = Self::default();
        for lvl in 0..9 {
            let tag = format!("lvl{}pPr", lvl + 1);
            let Some(lp) = list_style
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == tag)
            else {
                continue;
            };
            out.before[lvl] = paragraph_spacing(lp, "spcBef");
            out.after[lvl] = paragraph_spacing(lp, "spcAft");
            out.line[lvl] = child(lp, "lnSpc").and_then(parse_lnspc);
        }
        out
    }

    /// Per level and property, `self` where set, else `fallback`.
    pub(crate) fn or(&self, fallback: &Self) -> Self {
        let mut out = self.clone();
        for lvl in 0..9 {
            out.before[lvl] = out.before[lvl].or(fallback.before[lvl]);
            out.after[lvl] = out.after[lvl].or(fallback.after[lvl]);
            out.line[lvl] = out.line[lvl].take().or_else(|| fallback.line[lvl].clone());
        }
        out
    }

    pub(crate) fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

/// marL (EMU) of each level of PowerPoint's default presentation
/// `defaultTextStyle`: 0.5" per level. Observed (#1620 controls): in a deck
/// without a defaultTextStyle a level-2 text box paragraph started 36 pt in.
pub(crate) const DEFAULT_TEXT_STYLE_MAR_L: [i64; 9] = [
    0, 457_200, 914_400, 1_371_600, 1_828_800, 2_286_000, 2_743_200, 3_200_400, 3_657_600,
];

/// Per-list-level default font sizes (pt). Index 0..=8 → lvl1pPr..lvl9pPr
/// (ECMA-376 §21.1.2.4). `None` where the level isn't specified.
pub(crate) type LevelFontSizes = [Option<f64>; 9];

/// Read `<a:lvlNpPr><a:defRPr@sz>` for levels 1..9 from a node that holds
/// `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master `<p:txStyles>`
/// style node (`<p:bodyStyle>` etc.). Sizes are in pt.
pub(crate) fn read_level_font_sizes(list_style: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        *slot = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| attr_f64(&rp, "sz"))
            .map(|v| v / 100.0);
    }
    out
}

/// Per-level default font sizes from a txBody's own `<a:lstStyle>`.
pub(crate) fn extract_level_font_sizes(tx_body: roxmltree::Node<'_, '_>) -> LevelFontSizes {
    child(tx_body, "lstStyle")
        .map(read_level_font_sizes)
        .unwrap_or([None; 9])
}

/// True when any level carries a size (avoids storing all-None arrays).
pub(crate) fn has_any_level_size(s: &LevelFontSizes) -> bool {
    s.iter().any(|v| v.is_some())
}

/// Per-edge merge: `primary[lvl]` wins, else `fallback[lvl]`.
pub(crate) fn merge_level_sizes(
    primary: &LevelFontSizes,
    fallback: &LevelFontSizes,
) -> LevelFontSizes {
    let mut out: LevelFontSizes = [None; 9];
    for lvl in 0..9 {
        out[lvl] = primary[lvl].or(fallback[lvl]);
    }
    out
}

/// Typeface PowerPoint uses when no tier of the list-style chain names a Latin
/// face. Observed with PowerPoint for Mac PDF export (issue #1620): a style
/// level that is present without `<a:latin>`, a level that is absent, a theme
/// font slot that is empty, and a placeholder cut off from the master by a
/// layout slot without a txBody all render in Arial. It is not the theme minor
/// font: those decks had Calibri, Verdana, Tw Cen MT or Bookman themes.
pub(crate) const HARD_DEFAULT_LATIN_FACE: &str = "Arial";
/// Size (pt) under the same conditions as [`HARD_DEFAULT_LATIN_FACE`].
pub(crate) const HARD_DEFAULT_FONT_SIZE: f64 = 18.0;

/// Per-list-level Latin typefaces. Index 0..=8 maps to `lvl1pPr`..`lvl9pPr`.
/// Each level is independent: PowerPoint does not reuse the level-1 face for a
/// deeper level whose style omits `<a:latin>` (issue #1620 controls).
pub(crate) type LevelFaces = [Option<String>; 9];

/// Resolve one authored `<a:latin typeface>` value.
///
/// * A theme token (`+mj-lt`, `+mn-lt`, …, ECMA-376 §20.1.4.1.16-.17) is
///   checked against the slide master's own theme. A token naming an absent or
///   empty theme slot counts as unspecified (`None`), so the next tier of the
///   chain applies: with an empty theme minor font a text-box run carrying
///   `+mn-lt` took the presentation `defaultTextStyle` face, and a placeholder
///   took the hard default. A token naming a face is kept as the token: the
///   face depends on the run language (a theme script font such as `Viet`,
///   issue #1627) and is chosen by [`finalize_latin_face`] once the run's
///   cascade is complete.
/// * A literal empty `typeface=""` is an authored face that no installed font
///   matches. PowerPoint rendered it in Arial in a run, in master txStyles and
///   in a shape lstStyle, never falling through to the inherited face.
pub(crate) fn resolve_latin_face(
    typeface: &str,
    theme: &HashMap<String, String>,
) -> Option<String> {
    if typeface.is_empty() {
        return Some(HARD_DEFAULT_LATIN_FACE.to_owned());
    }
    if typeface.starts_with('+') {
        return theme
            .get(typeface)
            .filter(|face| !face.is_empty())
            .map(|_| typeface.to_owned());
    }
    Some(typeface.to_owned())
}

/// The face a resolved Latin chain value names for a run language: a kept
/// theme token picks its collection's script font for the language (Viet),
/// else the collection's Latin face. Literal faces pass through.
pub(crate) fn finalize_latin_face(
    face: Option<&str>,
    theme: &HashMap<String, String>,
    lang: Option<&str>,
) -> Option<String> {
    let face = face?;
    if theme_token_set(face).is_none() {
        return Some(face.to_owned());
    }
    resolve_slot_face(face, FontSlot::Latin, theme, lang, None)
}

/// The resolved Latin face of a `defRPr` / `rPr`, or `None` when unspecified.
pub(crate) fn run_properties_latin_face(
    properties: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    child(properties, "latin")
        .and_then(|latin| attr(&latin, "typeface"))
        .and_then(|face| resolve_latin_face(&face, theme))
}

/// Read `<a:lvlNpPr><a:defRPr><a:latin>` for levels 1..9 from a list-style
/// node (a txBody `<a:lstStyle>`, a master txStyles style, or the presentation
/// `<p:defaultTextStyle>`).
pub(crate) fn read_level_faces(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> LevelFaces {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| run_properties_latin_face(rp, theme))
    })
}

/// Per-level faces from a txBody's own `<a:lstStyle>`.
pub(crate) fn extract_level_faces(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> LevelFaces {
    child(tx_body, "lstStyle")
        .map(|ls| read_level_faces(ls, theme))
        .unwrap_or_default()
}

/// Per-level authored value of `<a:lvlNpPr><a:defRPr>`: the typeface of child
/// `element` (`ea` / `cs`, kept as authored) or, with `element` None, the
/// attribute `attribute` (`lang` / `altLang`).
pub(crate) fn read_level_defrpr_values(
    list_style: roxmltree::Node<'_, '_>,
    element: Option<&str>,
    attribute: &str,
) -> LevelFaces {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        let def_rpr = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))?;
        match element {
            Some(name) => child(def_rpr, name).and_then(|n| attr(&n, attribute)),
            None => attr(&def_rpr, attribute),
        }
    })
}

pub(crate) fn has_any_level_face(faces: &LevelFaces) -> bool {
    faces.iter().any(Option::is_some)
}

/// Per-level merge: `primary[lvl]` wins, else `fallback[lvl]`.
pub(crate) fn merge_level_faces(primary: &LevelFaces, fallback: &LevelFaces) -> LevelFaces {
    std::array::from_fn(|lvl| primary[lvl].clone().or_else(|| fallback[lvl].clone()))
}

/// End a face chain at the hard default so every level names a face.
pub(crate) fn complete_level_faces(faces: &LevelFaces) -> LevelFaces {
    std::array::from_fn(|lvl| {
        Some(
            faces[lvl]
                .clone()
                .unwrap_or_else(|| HARD_DEFAULT_LATIN_FACE.to_owned()),
        )
    })
}

/// End a size chain at the hard default so every level has a size.
pub(crate) fn complete_level_sizes(sizes: &LevelFontSizes) -> LevelFontSizes {
    std::array::from_fn(|lvl| Some(sizes[lvl].unwrap_or(HARD_DEFAULT_FONT_SIZE)))
}

/// Per-list-level paragraph alignment (`<a:lvlNpPr@algn>`). Index 0..=8 →
/// lvl1pPr..lvl9pPr; `None` where the level does not set it.
pub(crate) type LevelAlignments = [Option<String>; 9];

/// Read `<a:lvlNpPr@algn>` for levels 1..9 from a node holding `<a:lvlNpPr>`
/// children (a txBody's `<a:lstStyle>` or a master `<p:txStyles>` style).
pub(crate) fn read_level_alignments(list_style: roxmltree::Node<'_, '_>) -> LevelAlignments {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| attr(&lp, "algn"))
    })
}

/// Per-list-level default text colours. Index 0..=8 maps to
/// `lvl1pPr`..`lvl9pPr` (ECMA-376 §21.1.2.4). Keeping these per level is
/// essential: applying a layout's lvl1 colour as a body-wide fallback makes a
/// `pPr@lvl="1"` paragraph inherit the wrong tier.
pub(crate) type LevelColors = [Option<String>; 9];

pub(crate) fn read_level_colors(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> LevelColors {
    let mut out: LevelColors = std::array::from_fn(|_| None);
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        *slot = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| text_property_color(rp, theme));
    }
    out
}

pub(crate) fn extract_level_colors(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> LevelColors {
    child(tx_body, "lstStyle")
        .map(|list_style| read_level_colors(list_style, theme))
        .unwrap_or_else(|| std::array::from_fn(|_| None))
}

pub(crate) fn has_any_level_color(colors: &LevelColors) -> bool {
    colors.iter().any(Option::is_some)
}

pub(crate) fn merge_level_colors(primary: &LevelColors, fallback: &LevelColors) -> LevelColors {
    std::array::from_fn(|lvl| primary[lvl].clone().or_else(|| fallback[lvl].clone()))
}

/// Per-list-level paragraph indents (EMU) — the `marL`/`marR`/`indent` attributes
/// of a `<a:lvlNpPr>` (ECMA-376 §21.1.2.4.13; `lvlNpPr` is a
/// `CT_TextParagraphProperties`, so these are attributes ON the level element
/// itself, exactly like a paragraph's own `<a:pPr>`). Each axis is `Option` so it
/// inherits independently: a level that sets only `marL` leaves `marR`/`indent`
/// `None` and a lower-priority tier supplies them.
#[derive(Clone, Copy, Debug, Default, serde::Serialize)]
pub(crate) struct LevelIndent {
    pub(crate) mar_l: Option<i64>,
    pub(crate) mar_r: Option<i64>,
    pub(crate) indent: Option<i64>,
}
pub(crate) type LevelIndents = [LevelIndent; 9];

/// Read `<a:lvlNpPr@marL/@marR/@indent>` (EMU) for levels 1..9 from a node that
/// holds `<a:lvlNpPr>` children — a txBody's `<a:lstStyle>` or a master
/// `<p:txStyles>` style node. Mirrors `read_level_font_sizes`, but the values are
/// attributes of the `lvlNpPr` element itself (not of a `<a:defRPr>` child).
pub(crate) fn read_level_indents(list_style: roxmltree::Node<'_, '_>) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for (lvl, slot) in out.iter_mut().enumerate() {
        let tag = format!("lvl{}pPr", lvl + 1);
        if let Some(lp) = list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
        {
            slot.mar_l = attr_i64(&lp, "marL");
            slot.mar_r = attr_i64(&lp, "marR");
            slot.indent = attr_i64(&lp, "indent");
        }
    }
    out
}

/// Per-level indents from a txBody's own `<a:lstStyle>`.
pub(crate) fn extract_level_indents(tx_body: roxmltree::Node<'_, '_>) -> LevelIndents {
    child(tx_body, "lstStyle")
        .map(read_level_indents)
        .unwrap_or_default()
}

/// True when any level carries any indent axis (avoids storing all-None arrays).
pub(crate) fn has_any_level_indent(s: &LevelIndents) -> bool {
    s.iter()
        .any(|li| li.mar_l.is_some() || li.mar_r.is_some() || li.indent.is_some())
}

/// Per-level, per-axis merge: `primary[lvl].x` wins, else `fallback[lvl].x`.
pub(crate) fn merge_level_indents(primary: &LevelIndents, fallback: &LevelIndents) -> LevelIndents {
    let mut out: LevelIndents = Default::default();
    for lvl in 0..9 {
        out[lvl].mar_l = primary[lvl].mar_l.or(fallback[lvl].mar_l);
        out[lvl].mar_r = primary[lvl].mar_r.or(fallback[lvl].mar_r);
        out[lvl].indent = primary[lvl].indent.or(fallback[lvl].indent);
    }
    out
}

/// The marker choice of a bullet (ECMA-376 §21.1.2.4 EG_TextBullet:
/// `buNone`/`buAutoNum`/`buChar`/`buBlip`). Separate from the three decoration
/// groups (colour/size/typeface) so each inherits independently across the
/// style cascade — see [`BulletProps`].
#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub(crate) enum BuMarker {
    /// `<a:buNone>` — explicitly no marker (§21.1.2.4.8).
    None,
    /// `<a:buChar char="…">` (§21.1.2.4.3).
    Char(String),
    /// `<a:buAutoNum type startAt>` (§21.1.2.4.1).
    AutoNum {
        num_type: String,
        start_at: Option<u32>,
    },
    /// `<a:buBlip>` resolved to an embedded zip path + mime (§21.1.2.4.2).
    Blip {
        image_path: String,
        mime_type: String,
    },
}

/// Bullet colour group (ECMA-376 §21.1.2.4 EG_TextBulletColor): an explicit
/// `<a:buClr>` colour or `<a:buClrTx>` "follow the run's text colour". Absence
/// is modelled by the enclosing `Option` (inherit from a lower style tier).
#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub(crate) enum BuColor {
    /// `<a:buClrTx>` (§21.1.2.4.5) — follow the text run's colour.
    FollowText,
    /// `<a:buClr>` (§21.1.2.4.4) — explicit resolved colour (hex, no '#').
    Color(String),
}

/// Bullet typeface group (EG_TextBulletTypeface): `<a:buFont>` explicit or
/// `<a:buFontTx>` follow-text. Absence = inherit (enclosing `Option`).
#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub(crate) enum BuFont {
    /// `<a:buFontTx>` (§21.1.2.4.7) — follow the text run's font.
    FollowText,
    /// `<a:buFont typeface>` (§21.1.2.4.6) — explicit resolved typeface.
    Font(String),
}

/// Bullet size group (EG_TextBulletSize): `<a:buSzTx>` follow-text,
/// `<a:buSzPct>` percent-of-text, or `<a:buSzPts>` absolute point size. Absence =
/// inherit (enclosing `Option`). The three are members of one `xsd:choice`, so at
/// most one is present at a level; whichever is present BLOCKS a lower-tier size.
///
/// `Pct` and `Pts` are separate variants (never both) because they collapse to
/// distinct resolved fields — `sizePct` (percent of the run size, resolved at
/// draw time) and `sizePts` (absolute points, run-independent) — on the [`Bullet`]
/// contract. Keeping them apart lets the renderer honour absolute-point bullets
/// (§21.1.2.4.10) while the cascade still treats the whole size group uniformly.
#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub(crate) enum BuSize {
    /// `<a:buSzTx>` (§21.1.2.4.11) — follow the text run's size.
    FollowText,
    /// `<a:buSzPct val>` (§21.1.2.4.9) — percentage of text size (100.0 = 100%).
    Pct(f64),
    /// `<a:buSzPts val>` (§21.1.2.4.10) — absolute size in points (18.0 = 18pt).
    Pts(f64),
}

/// A paragraph's bullet as FOUR independent choice groups (ECMA-376 §21.1.2.4:
/// `CT_TextParagraphProperties` carries EG_TextBulletColor, EG_TextBulletSize,
/// EG_TextBulletTypeface and EG_TextBullet as separate optional children). Each
/// group inherits per-property across the master → layout → txBody-lstStyle →
/// paragraph-pPr cascade, so a decoration declared in one tier survives onto a
/// marker declared in another (PowerPoint resolves each group independently).
/// `None` on a field means "not specified here — inherit from a lower tier";
/// this is the state that whole-`Bullet` merging could not represent, which is
/// why cross-tier splits (e.g. master `buAutoNum` + slide `buClr`) lost the
/// higher-priority colour/size/font. Collapsed into the serialized [`Bullet`]
/// contract at the leaf via [`BulletProps::resolve`].
#[derive(Clone, Debug, Default, PartialEq, serde::Serialize)]
pub(crate) struct BulletProps {
    pub(crate) marker: Option<BuMarker>,
    pub(crate) color: Option<BuColor>,
    pub(crate) font: Option<BuFont>,
    pub(crate) size: Option<BuSize>,
}

impl BulletProps {
    /// True when no group is specified (all inherit). A level with only a
    /// decoration (e.g. `buClr` but no marker) is NOT inherit — it must survive
    /// the cascade to reach an inherited marker.
    pub(crate) fn is_inherit(&self) -> bool {
        self.marker.is_none() && self.color.is_none() && self.font.is_none() && self.size.is_none()
    }

    /// Field-wise merge: for each group `primary` wins when it specifies that
    /// group, else `fallback` supplies it. An explicit `*Tx` (follow-text) value
    /// counts as "specified" and therefore BLOCKS a lower tier — that is the
    /// whole point of e.g. `buClrTx` over an inherited `buClr`.
    pub(crate) fn merge(primary: &BulletProps, fallback: &BulletProps) -> BulletProps {
        BulletProps {
            marker: primary.marker.clone().or_else(|| fallback.marker.clone()),
            color: primary.color.clone().or_else(|| fallback.color.clone()),
            font: primary.font.clone().or_else(|| fallback.font.clone()),
            size: primary.size.clone().or_else(|| fallback.size.clone()),
        }
    }

    /// Collapse the resolved groups into the serialized [`Bullet`] the renderer
    /// consumes. Collapsing happens ONLY at the leaf (never mid-cascade), so the
    /// follow-text-vs-inherit distinction is preserved while merging:
    /// - colour: `Color(c)` → `Some(c)`; `FollowText`/inherit → `None` (the
    ///   renderer already treats `None` as "follow the run's colour");
    /// - size: `Pct(p)` → `size_pct = Some(p)`; `Pts(p)` → `size_pts = Some(p)`
    ///   (§21.1.2.4.10, absolute points); `FollowText`/inherit → both `None`. The
    ///   two size fields are mutually exclusive (one `xsd:choice`);
    /// - font: `Font(f)` → `Some(f)`; `FollowText`/inherit → `None`.
    ///   Auto-number markers retain both font and size because their glyph
    ///   metrics determine whether multi-digit labels fit the hanging gutter.
    ///   Picture markers retain size but have no applicable font.
    pub(crate) fn resolve(&self) -> Bullet {
        let color = match &self.color {
            Some(BuColor::Color(c)) => Some(c.clone()),
            _ => None,
        };
        let size_pct = match &self.size {
            Some(BuSize::Pct(v)) => Some(*v),
            _ => None,
        };
        // Absolute-point size (§21.1.2.4.10). Exclusive with `size_pct` above (one
        // `xsd:choice`), so at most one of the two is `Some` on a resolved bullet.
        let size_pts = match &self.size {
            Some(BuSize::Pts(v)) => Some(*v),
            _ => None,
        };
        let font_family = match &self.font {
            Some(BuFont::Font(f)) => Some(f.clone()),
            _ => None,
        };
        match &self.marker {
            None => Bullet::Inherit,
            Some(BuMarker::None) => Bullet::None,
            Some(BuMarker::Char(ch)) => Bullet::Char {
                ch: ch.clone(),
                color,
                size_pct,
                size_pts,
                font_family,
            },
            Some(BuMarker::AutoNum { num_type, start_at }) => Bullet::AutoNum {
                num_type: num_type.clone(),
                start_at: *start_at,
                color,
                size_pct,
                size_pts,
                font_family,
            },
            Some(BuMarker::Blip {
                image_path,
                mime_type,
            }) => Bullet::Blip {
                image_path: image_path.clone(),
                mime_type: mime_type.clone(),
                size_pct,
                size_pts,
            },
        }
    }
}

/// Per-list-level bullet properties (index 0..=8 → lvl1pPr..lvl9pPr). Each level
/// is a [`BulletProps`] whose groups are individually `None` when the level's
/// `<a:lvlNpPr>` (or the paragraph's `<a:pPr>`) does not specify them, so lower
/// style tiers can supply the missing groups per-property.
pub(crate) type LevelBullets = [BulletProps; 9];

pub(crate) fn empty_level_bullets() -> LevelBullets {
    std::array::from_fn(|_| BulletProps::default())
}

/// True when any level specifies any bullet group (avoids storing all-inherit
/// arrays). Must be decoration-aware: a level carrying only a `buClr` (no
/// marker) still has to be stored so the colour reaches an inherited marker.
pub(crate) fn has_any_level_bullet(s: &LevelBullets) -> bool {
    s.iter().any(|b| !b.is_inherit())
}

/// Read `<a:lvlNpPr>` bullet groups for levels 1..9 from a node holding
/// `<a:lvlNpPr>` children (a txBody `<a:lstStyle>` or a master `<p:txStyles>`
/// style node). Each level captures whichever of the four choice groups it
/// declares; absent groups stay `None` so lower tiers supply them.
pub(crate) fn read_level_bullets<F: FnMut(&str) -> Option<String>>(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    std::array::from_fn(|lvl| {
        let tag = format!("lvl{}pPr", lvl + 1);
        list_style
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == tag)
            .map(|lp| parse_bullet_props(Some(lp), theme, resolve_blip))
            .unwrap_or_default()
    })
}

/// Per-level bullets from a txBody's own `<a:lstStyle>`. `resolve_blip` resolves
/// a level's `<a:buBlip>` embed against this text body's part rels (§21.1.2.4.2).
pub(crate) fn extract_level_bullets<F: FnMut(&str) -> Option<String>>(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> LevelBullets {
    child(tx_body, "lstStyle")
        .map(|ls| read_level_bullets(ls, theme, resolve_blip))
        .unwrap_or_else(empty_level_bullets)
}

/// Per-level, per-group merge: `primary[lvl]` wins each group it specifies, else
/// `fallback[lvl]` supplies it (see [`BulletProps::merge`]).
pub(crate) fn merge_level_bullets(primary: &LevelBullets, fallback: &LevelBullets) -> LevelBullets {
    std::array::from_fn(|lvl| BulletProps::merge(&primary[lvl], &fallback[lvl]))
}

// ===========================
//  Text body parsing
// ===========================

/// Return the text-property fill choice only when it appears in the
/// `CT_TextCharacterProperties` sequence position defined by ECMA-376
/// §21.1.2.3.9 / dml-main.xsd. The fill choice precedes effects, highlight,
/// underline properties and the latin/ea/cs font children. PowerPoint ignores
/// an out-of-order fill (a pattern emitted by some non-Office producers), so a
/// name-only descendant lookup would invent formatting Office does not apply.
fn text_property_fill<'a, 'input>(
    properties: roxmltree::Node<'a, 'input>,
) -> Option<roxmltree::Node<'a, 'input>> {
    let sequence_rank = |name: &str| -> Option<u8> {
        match name {
            "ln" => Some(0),
            "noFill" | "solidFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill" => Some(1),
            "effectLst" | "effectDag" => Some(2),
            "highlight" => Some(3),
            "uLnTx" | "uLn" => Some(4),
            "uFillTx" | "uFill" => Some(5),
            "latin" => Some(6),
            "ea" => Some(7),
            "cs" => Some(8),
            "sym" => Some(9),
            "hlinkClick" => Some(10),
            "hlinkMouseOver" => Some(11),
            "rtl" => Some(12),
            "extLst" => Some(13),
            _ => None,
        }
    };
    let mut highest_preceding_rank = 0_u8;
    for node in properties.children().filter(|node| node.is_element()) {
        let name = node.tag_name().name();
        let Some(rank) = sequence_rank(name) else {
            continue;
        };
        if matches!(
            name,
            "noFill" | "solidFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill"
        ) {
            return (highest_preceding_rank <= rank).then_some(node);
        }
        highest_preceding_rank = highest_preceding_rank.max(rank);
    }
    None
}

/// Resolve the colour of a text fill when it is either a solid fill or a
/// gradient whose every stop has the same resolved colour. The latter is
/// visually a solid colour despite its gradient encoding, so it fits the
/// existing text-run colour model without approximating a genuine gradient.
pub(crate) fn text_property_color(
    properties: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    let fill = text_property_fill(properties)?;
    match fill.tag_name().name() {
        "solidFill" => parse_color_node(fill, theme),
        "gradFill" => {
            let colors = child(fill, "gsLst")?
                .children()
                .filter(|node| node.is_element() && node.tag_name().name() == "gs")
                .map(|stop| parse_color_node(stop, theme))
                .collect::<Option<Vec<_>>>()?;
            let first = colors.first()?.clone();
            colors
                .into_iter()
                .all(|color| color == first)
                .then_some(first)
        }
        _ => None,
    }
}

fn underline_fill_choice<'a, 'input>(
    properties: roxmltree::Node<'a, 'input>,
) -> Option<roxmltree::Node<'a, 'input>> {
    properties
        .children()
        .find(|node| node.is_element() && matches!(node.tag_name().name(), "uFill" | "uFillTx"))
}

/// A `<a:prstTxWarp>` child of `<a:bodyPr>`: `None` when absent, `Some(None)`
/// for an explicit `textNoShape` (no warp), `Some(Some(_))` for a warp preset.
fn parse_text_warp(body_pr: roxmltree::Node<'_, '_>) -> Option<Option<TextWarp>> {
    let warp = child(body_pr, "prstTxWarp")?;
    // ST_TextShapeType derives from xsd:token (whiteSpace="collapse"), so the
    // surrounding XML whitespace (#x20 #x9 #xD #xA) is not part of the value.
    // Strip only that: no Unicode-whitespace trim, no internal collapse, no
    // case folding: every other codepoint and spelling stays as authored.
    let raw = attr(&warp, "prst").unwrap_or_default();
    let preset = raw
        .trim_matches(|c: char| matches!(c, ' ' | '\t' | '\r' | '\n'))
        .to_string();
    if preset.is_empty() || preset == "textNoShape" {
        return Some(None);
    }
    let adj = child(warp, "avLst")
        .map(|av| {
            av.children()
                .filter(|c| c.is_element() && c.tag_name().name() == "gd")
                .filter_map(|gd| {
                    // fmla is "val <n>" for avLst adjust guides.
                    attr(&gd, "fmla").and_then(|f| {
                        f.strip_prefix("val ")
                            .and_then(|v| v.trim().parse::<i64>().ok())
                    })
                })
                .collect::<Vec<i64>>()
        })
        .unwrap_or_default();
    Some(Some(TextWarp { preset, adj }))
}

#[cfg(test)]
mod prst_tx_warp_token_tests {
    use super::parse_text_warp;

    /// `None` = no prstTxWarp (inherit), `Some(None)` = explicit no warp
    /// (overrides an inherited warp), `Some(Some(p))` = warp preset `p`.
    fn warp_of(inner: &str) -> Option<Option<String>> {
        let xml = format!(
            r#"<bodyPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">{inner}</bodyPr>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let body_pr = doc.root_element();
        parse_text_warp(body_pr).map(|w| w.map(|w| w.preset))
    }

    #[test]
    fn padded_no_shape_stays_an_explicit_override_not_absence() {
        assert_eq!(warp_of(""), None);
        assert_eq!(warp_of(r#"<prstTxWarp prst="textNoShape"/>"#), Some(None));
        assert_eq!(
            warp_of(r#"<prstTxWarp prst=" textNoShape&#9;"/>"#),
            Some(None)
        );
        assert_eq!(
            warp_of(r#"<prstTxWarp prst="&#13;&#10;textNoShape "/>"#),
            Some(None)
        );
        // Whitespace-only collapses to the empty token, handled like prst="".
        assert_eq!(warp_of(r#"<prstTxWarp prst=" &#9; "/>"#), Some(None));
        // Non-XML whitespace (NBSP, EM SPACE) is not token padding.
        assert_eq!(
            warp_of(r#"<prstTxWarp prst="&#160;textNoShape&#x2003;"/>"#),
            Some(Some("\u{a0}textNoShape\u{2003}".to_string()))
        );
    }
}

/// Stored autofit child (`<a:spAutoFit>` / `<a:normAutofit>` / `<a:noAutofit>`).
#[derive(Clone, Debug, serde::Serialize)]
pub(crate) struct InheritedAutoFit {
    pub(crate) mode: String,
    pub(crate) font_scale: Option<f64>,
    pub(crate) ln_spc_reduction: Option<f64>,
}

/// The `<a:bodyPr>` values a placeholder inherits from its layout and master
/// placeholder (ECMA-376 §19.3.1.36 placeholder inheritance). Each field is
/// `None` when that level omits it, so the cascade can fall through
/// attribute by attribute: slide → layout → master → schema default.
///
/// Observed with PowerPoint for Mac PDF export (issue #1618): for wrap, vert,
/// numCol, spcCol, rtlCol, spcFirstLastPara, prstTxWarp and normAutofit with
/// a stored fontScale, a slide placeholder that omits the value takes the
/// layout's, else the master's, else the schema default, and the slide's own
/// value always wins — 256/256 cases covering master-only, layout-only,
/// both (layout wins, including an explicit default on the layout) and
/// neither, each with and without a conflicting theme objectDefaults value
/// (never used). Anchor keeps its own idx/type lookup (`lookup_anchor`).
#[derive(Clone, Debug, Default, serde::Serialize)]
pub(crate) struct InheritedBodyPr {
    /// `lIns`, `tIns`, `rIns`, `bIns` (EMU).
    pub(crate) insets: [Option<i64>; 4],
    pub(crate) wrap: Option<String>,
    pub(crate) vert: Option<String>,
    pub(crate) num_col: Option<u32>,
    pub(crate) spc_col: Option<i64>,
    pub(crate) rtl_col: Option<bool>,
    pub(crate) spc_first_last_para: Option<bool>,
    /// `anchorCtr` (ECMA-376 §21.1.2.1.1).
    pub(crate) anchor_ctr: Option<bool>,
    /// `compatLnSpc` (see the cascade note on `TextBody::compat_ln_spc`).
    pub(crate) compat_ln_spc: Option<bool>,
    pub(crate) auto_fit: Option<InheritedAutoFit>,
    pub(crate) text_warp: Option<Option<TextWarp>>,
}

impl InheritedBodyPr {
    pub(crate) fn from_body_pr(body_pr: roxmltree::Node<'_, '_>) -> Self {
        let flag = |name: &str| attr(&body_pr, name).map(|v| v == "1" || v == "true");
        Self {
            insets: [
                attr_i64(&body_pr, "lIns"),
                attr_i64(&body_pr, "tIns"),
                attr_i64(&body_pr, "rIns"),
                attr_i64(&body_pr, "bIns"),
            ],
            wrap: attr(&body_pr, "wrap"),
            vert: attr(&body_pr, "vert"),
            num_col: attr(&body_pr, "numCol")
                .and_then(|v| v.parse::<u32>().ok())
                .filter(|&n| n >= 1),
            spc_col: attr_i64(&body_pr, "spcCol"),
            rtl_col: flag("rtlCol"),
            spc_first_last_para: flag("spcFirstLastPara"),
            anchor_ctr: flag("anchorCtr"),
            compat_ln_spc: flag("compatLnSpc"),
            auto_fit: ooxml_common::text::parse_autofit(body_pr).map(
                |(mode, font_scale, ln_spc_reduction)| InheritedAutoFit {
                    mode,
                    font_scale,
                    ln_spc_reduction,
                },
            ),
            text_warp: parse_text_warp(body_pr),
        }
    }

    /// Field-wise `self.or(fallback)`: this level's value wins where present.
    pub(crate) fn or(self, fallback: &Self) -> Self {
        Self {
            insets: std::array::from_fn(|i| self.insets[i].or(fallback.insets[i])),
            wrap: self.wrap.or_else(|| fallback.wrap.clone()),
            vert: self.vert.or_else(|| fallback.vert.clone()),
            num_col: self.num_col.or(fallback.num_col),
            spc_col: self.spc_col.or(fallback.spc_col),
            rtl_col: self.rtl_col.or(fallback.rtl_col),
            spc_first_last_para: self.spc_first_last_para.or(fallback.spc_first_last_para),
            anchor_ctr: self.anchor_ctr.or(fallback.anchor_ctr),
            compat_ln_spc: self.compat_ln_spc.or(fallback.compat_ln_spc),
            auto_fit: self.auto_fit.or_else(|| fallback.auto_fit.clone()),
            text_warp: self.text_warp.or_else(|| fallback.text_warp.clone()),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.insets.iter().all(Option::is_none)
            && self.wrap.is_none()
            && self.vert.is_none()
            && self.num_col.is_none()
            && self.spc_col.is_none()
            && self.rtl_col.is_none()
            && self.spc_first_last_para.is_none()
            && self.anchor_ctr.is_none()
            && self.compat_ln_spc.is_none()
            && self.auto_fit.is_none()
            && self.text_warp.is_none()
    }
}

fn underline_line_choice<'a, 'input>(
    properties: roxmltree::Node<'a, 'input>,
) -> Option<roxmltree::Node<'a, 'input>> {
    properties
        .children()
        .find(|node| node.is_element() && matches!(node.tag_name().name(), "uLn" | "uLnTx"))
}

fn parse_text_line(
    line: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<TextOutline> {
    if child(line, "noFill").is_some() {
        return None;
    }
    let fill = parse_fill(line, theme);
    Some(TextOutline {
        width: attr_i64(&line, "w").unwrap_or(0),
        color: match &fill {
            Some(Fill::Solid { color }) => Some(color.clone()),
            _ => None,
        },
        fill,
    })
}

/// The authored part of CT_TextCharacterProperties. Each member retains its
/// own presence bit so a partial pPr/defRPr overrides only that property of
/// lstStyle, layout and master (§21.1.2.2.7, §21.1.2.4, §21.1.2.3.9).
/// An explicit noFill is stored as Fill::None and never becomes "missing".
#[derive(Clone, Default, serde::Serialize)]
pub(crate) struct RunProperties {
    /// Authored CT_TextCharacterProperties attributes, except dirty/err (the
    /// schema marks these as editing diagnostics).  Keep the full set so
    /// language, kerning, normalization and annotation state survive the same
    /// cascade as visible paint even when this canvas has no consumer yet.
    attributes: PropertyAttributes,
    /// Direct child attributes, merged within a child after schema choice
    /// groups have selected the nearest authored alternative.
    child_attributes: BTreeMap<String, PropertyAttributes>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<String>,
    underline_fill: Option<Option<Fill>>,
    underline_line: Option<Option<TextOutline>>,
    underline_line_follow_text: Option<bool>,
    underline_line_fill_authored: bool,
    strike: Option<String>,
    caps: Option<String>,
    letter_spacing: Option<f64>,
    font_size: Option<f64>,
    fill: Option<Fill>,
    color: Option<String>,
    fill_authored: bool,
    font_family: Option<String>,
    font_family_ea: Option<String>,
    font_family_cs: Option<String>,
    font_family_sym: Option<String>,
    baseline: Option<i32>,
    effects_authored: bool,
    shadow: Option<Shadow>,
    reflection: Option<Reflection>,
    outline: Option<Option<TextOutline>>,
    outline_fill_authored: bool,
    highlight: Option<String>,
    hyperlink_uses_text_fill: Option<bool>,
    // Relationship IDs are local to their owning OPC part. Keep the target
    // paired with the authored r:id as levels from master, layout and slide
    // are merged (ECMA-376 Part 2, §9.3.3; §21.1.2.3.5).
    hlink_click_target: Option<Option<InheritedRelationship>>,
    hlink_mouse_over_target: Option<Option<InheritedRelationship>>,
}
pub(crate) type LevelRunProperties = [RunProperties; 9];

impl RunProperties {
    pub(crate) fn from_xml(node: roxmltree::Node<'_, '_>, theme: &HashMap<String, String>) -> Self {
        let fill_choice = text_property_fill(node);
        let underline_choice = underline_fill_choice(node);
        let underline_line_choice = underline_line_choice(node);
        let effects = child(node, "effectLst").or_else(|| child(node, "effectDag"));
        let outline = child(node, "ln");
        let mut child_attributes = BTreeMap::new();
        for element in node.children().filter(|n| n.is_element()) {
            child_attributes.insert(
                element.tag_name().name().to_owned(),
                xml_attributes(element),
            );
        }
        Self {
            attributes: xml_attributes(node),
            child_attributes,
            bold: attr(&node, "b").map(|v| v == "1" || v == "true"),
            italic: attr(&node, "i").map(|v| v == "1" || v == "true"),
            underline: attr(&node, "u"),
            underline_fill: underline_choice.map(|n| {
                if n.tag_name().name() == "uFill" {
                    parse_fill(n, theme)
                } else {
                    None
                }
            }),
            underline_line: underline_line_choice.map(|n| {
                (n.tag_name().name() == "uLn")
                    .then(|| parse_text_line(n, theme))
                    .flatten()
            }),
            underline_line_follow_text: underline_line_choice
                .map(|n| n.tag_name().name() == "uLnTx"),
            underline_line_fill_authored: underline_line_choice
                .is_some_and(|n| n.tag_name().name() == "uLn" && text_property_fill(n).is_some()),
            strike: attr(&node, "strike"),
            caps: attr(&node, "cap"),
            letter_spacing: attr(&node, "spc").and_then(|v| text_point_to_pt(&v)),
            font_size: attr_f64(&node, "sz").map(|v| v / 100.0),
            fill: fill_choice.and_then(|_| parse_fill(node, theme)),
            color: fill_choice.and_then(|_| text_property_color(node, theme)),
            fill_authored: fill_choice.is_some(),
            // See `resolve_latin_face`: an empty theme slot is unspecified,
            // a literal empty typeface is Arial.
            font_family: child(node, "latin")
                .and_then(|n| attr(&n, "typeface"))
                .and_then(|v| resolve_latin_face(&v, theme)),
            // ea/cs keep the authored value (a token, a literal or "") so the
            // face can follow the run language after the cascade (#1627).
            font_family_ea: child(node, "ea").and_then(|n| attr(&n, "typeface")),
            font_family_cs: child(node, "cs").and_then(|n| attr(&n, "typeface")),
            font_family_sym: child(node, "sym")
                .and_then(|n| attr(&n, "typeface"))
                .map(|v| resolve_theme_typeface(&v, theme)),
            baseline: attr(&node, "baseline").and_then(|v| v.parse().ok()),
            effects_authored: effects.is_some(),
            shadow: effects.and_then(|n| parse_shadow(n, theme)),
            reflection: effects.and_then(parse_reflection),
            outline: outline.map(|ln| {
                if child(ln, "noFill").is_some() {
                    return None;
                }
                let fill = parse_fill(ln, theme);
                Some(TextOutline {
                    width: attr_i64(&ln, "w").unwrap_or(0),
                    color: match &fill {
                        Some(Fill::Solid { color }) => Some(color.clone()),
                        _ => None,
                    },
                    fill,
                })
            }),
            outline_fill_authored: outline.is_some_and(|ln| text_property_fill(ln).is_some()),
            highlight: child(node, "highlight").and_then(|n| parse_color_node(n, theme)),
            hyperlink_uses_text_fill: child(node, "hlinkClick")
                .and_then(|h| {
                    h.descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "hlinkClr")
                })
                .map(|n| n.attribute("val") == Some("tx")),
            hlink_click_target: None,
            hlink_mouse_over_target: None,
        }
    }

    pub(crate) fn with_relationships(mut self, rels: &HashMap<String, String>) -> Self {
        self.hlink_click_target = self
            .child_attributes
            .get("hlinkClick")
            .and_then(|attrs| attrs.get("id"))
            .map(|id| {
                rels.get(id).map(|target| InheritedRelationship {
                    target: target.clone(),
                    source_part: None,
                })
            });
        self.hlink_mouse_over_target = self
            .child_attributes
            .get("hlinkMouseOver")
            .and_then(|attrs| attrs.get("id"))
            .map(|id| {
                rels.get(id).map(|target| InheritedRelationship {
                    target: target.clone(),
                    source_part: None,
                })
            });
        self
    }

    /// Retain the relationship owner until after the attribute-wise cascade.
    /// A nearer level may author @action without a new r:id, so resolving the
    /// target here would miss a later hlinksldjump action.
    pub(crate) fn with_part_targets(mut self, source_part: &str) -> Self {
        for link in [
            &mut self.hlink_click_target,
            &mut self.hlink_mouse_over_target,
        ] {
            if let Some(Some(relationship)) = link {
                relationship.source_part = Some(source_part.to_owned());
            }
        }
        self
    }

    /// `self` has higher priority. Every field, including explicit false/none,
    /// is independently chosen; a present effect list is one OOXML choice.
    /// Replace the Latin face and size with the resolved list-style chain
    /// (placeholder or defaultTextStyle tiers, see shape.rs). The chain is
    /// authoritative for these two attributes; every other character
    /// property keeps its own cascade.
    pub(crate) fn with_chain_face_and_size(
        mut self,
        face: Option<String>,
        size: Option<f64>,
    ) -> Self {
        self.font_family = face;
        self.font_size = size;
        self
    }

    /// The ea/cs faces and language ordinary text inherits from the
    /// presentation defaultTextStyle level or its shape style's fontRef.
    pub(crate) fn script_base(
        ea: Option<String>,
        cs: Option<String>,
        lang: Option<String>,
        alt_lang: Option<String>,
    ) -> Self {
        let mut attributes = PropertyAttributes::new();
        if let Some(lang) = lang {
            attributes.insert("lang".to_owned(), lang);
        }
        if let Some(alt_lang) = alt_lang {
            attributes.insert("altLang".to_owned(), alt_lang);
        }
        Self {
            attributes,
            font_family_ea: ea,
            font_family_cs: cs,
            ..Default::default()
        }
    }

    /// The cascaded run language (`lang`, ECMA-376 §21.1.2.3.9).
    pub(crate) fn language(&self) -> Option<&str> {
        self.attributes
            .get("lang")
            .map(String::as_str)
            .filter(|v| !v.is_empty())
    }

    /// The cascaded alternate language (`altLang`).
    pub(crate) fn alt_language(&self) -> Option<&str> {
        self.attributes
            .get("altLang")
            .map(String::as_str)
            .filter(|v| !v.is_empty())
    }

    pub(crate) fn over(&self, lower: &Self) -> Self {
        macro_rules! pick {
            ($field:ident) => {
                self.$field.clone().or_else(|| lower.$field.clone())
            };
        }
        Self {
            attributes: merge_attributes(&self.attributes, &lower.attributes),
            child_attributes: {
                let mut merged = lower.child_attributes.clone();
                // The schema groups are alternatives.  A nearer choice ends
                // inheritance of the other choices, while attributes inside
                // the same child remain independently inherited.
                const FILL: &[&str] = &[
                    "noFill",
                    "solidFill",
                    "gradFill",
                    "blipFill",
                    "pattFill",
                    "grpFill",
                ];
                const EFFECT: &[&str] = &["effectLst", "effectDag"];
                const ULINE: &[&str] = &["uLn", "uLnTx"];
                const UFILL: &[&str] = &["uFill", "uFillTx"];
                for (name, attrs) in &self.child_attributes {
                    for group in [FILL, EFFECT, ULINE, UFILL] {
                        if group.contains(&name.as_str()) {
                            for other in group {
                                if *other != name {
                                    merged.remove(*other);
                                }
                            }
                        }
                    }
                    let previous = merged.get(name).cloned().unwrap_or_default();
                    merged.insert(name.clone(), merge_attributes(attrs, &previous));
                }
                merged
            },
            bold: pick!(bold),
            italic: pick!(italic),
            underline: pick!(underline),
            underline_fill: pick!(underline_fill),
            underline_line: match (&self.underline_line, &lower.underline_line) {
                (Some(Some(higher)), Some(Some(lower_line)))
                    if !self.underline_line_follow_text.unwrap_or(false) =>
                {
                    let fill = if self.underline_line_fill_authored {
                        higher.fill.clone()
                    } else {
                        lower_line.fill.clone()
                    };
                    Some(Some(TextOutline {
                        width: if self
                            .child_attributes
                            .get("uLn")
                            .is_some_and(|a| a.contains_key("w"))
                        {
                            higher.width
                        } else {
                            lower_line.width
                        },
                        color: match &fill {
                            Some(Fill::Solid { color }) => Some(color.clone()),
                            _ => None,
                        },
                        fill,
                    }))
                }
                (Some(value), _) => Some(value.clone()),
                (None, _) => lower.underline_line.clone(),
            },
            underline_line_follow_text: pick!(underline_line_follow_text),
            underline_line_fill_authored: self.underline_line_fill_authored
                || lower.underline_line_fill_authored,
            strike: pick!(strike),
            caps: pick!(caps),
            letter_spacing: pick!(letter_spacing),
            font_size: pick!(font_size),
            fill: if self.fill_authored {
                self.fill.clone()
            } else {
                lower.fill.clone()
            },
            color: if self.fill_authored {
                self.color.clone()
            } else {
                lower.color.clone()
            },
            fill_authored: self.fill_authored || lower.fill_authored,
            font_family: pick!(font_family),
            font_family_ea: pick!(font_family_ea),
            font_family_cs: pick!(font_family_cs),
            font_family_sym: pick!(font_family_sym),
            baseline: pick!(baseline),
            effects_authored: self.effects_authored || lower.effects_authored,
            shadow: if self.effects_authored {
                self.shadow.clone()
            } else {
                lower.shadow.clone()
            },
            reflection: if self.effects_authored {
                self.reflection.clone()
            } else {
                lower.reflection.clone()
            },
            outline: match (&self.outline, &lower.outline) {
                (Some(Some(higher)), Some(Some(lower_outline))) => {
                    let fill = if self.outline_fill_authored {
                        higher.fill.clone()
                    } else {
                        lower_outline.fill.clone()
                    };
                    Some(Some(TextOutline {
                        width: if self
                            .child_attributes
                            .get("ln")
                            .is_some_and(|a| a.contains_key("w"))
                        {
                            higher.width
                        } else {
                            lower_outline.width
                        },
                        color: match &fill {
                            Some(Fill::Solid { color }) => Some(color.clone()),
                            _ => None,
                        },
                        fill,
                    }))
                }
                (Some(value), _) => Some(value.clone()),
                (None, _) => lower.outline.clone(),
            },
            outline_fill_authored: self.outline_fill_authored || lower.outline_fill_authored,
            highlight: pick!(highlight),
            hyperlink_uses_text_fill: pick!(hyperlink_uses_text_fill),
            hlink_click_target: pick!(hlink_click_target),
            hlink_mouse_over_target: pick!(hlink_mouse_over_target),
        }
    }
    pub(crate) fn is_empty(&self) -> bool {
        self.attributes.is_empty()
            && self.child_attributes.is_empty()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.underline_fill.is_none()
            && self.underline_line.is_none()
            && self.strike.is_none()
            && self.caps.is_none()
            && self.letter_spacing.is_none()
            && self.font_size.is_none()
            && !self.fill_authored
            && self.font_family.is_none()
            && self.font_family_ea.is_none()
            && self.font_family_cs.is_none()
            && self.font_family_sym.is_none()
            && self.baseline.is_none()
            && !self.effects_authored
            && self.outline.is_none()
            && self.highlight.is_none()
            && self.hyperlink_uses_text_fill.is_none()
            && self.hlink_click_target.is_none()
            && self.hlink_mouse_over_target.is_none()
    }
    pub(crate) fn without_fill(mut self) -> Self {
        self.fill = None;
        self.color = None;
        self.fill_authored = false;
        self
    }
}

pub(crate) fn read_level_run_properties_with_rels(
    list_style: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> LevelRunProperties {
    // CT_TextListStyle.defPPr supplies the run defaults for every level.  A
    // level's defRPr overlays it one property at a time (§21.1.2.4).
    // Observed with PowerPoint for Mac PDF export (#1620 controls): a defPPr
    // Latin face or size had no effect in any tier — shape lstStyle, layout
    // slot, master placeholder, txStyles and defaultTextStyle — neither alone
    // nor under an lvlNpPr that omits it; the level fell through as if defPPr
    // were absent. Other defPPr character properties were not observable
    // there and keep the §21.1.2.4 base role.
    let base = child(list_style, "defPPr")
        .and_then(|p| child(p, "defRPr"))
        .map(|r| {
            RunProperties::from_xml(r, theme)
                .with_relationships(rels)
                .with_chain_face_and_size(None, None)
        })
        .unwrap_or_default();
    std::array::from_fn(|level| {
        child(list_style, &format!("lvl{}pPr", level + 1))
            .and_then(|p| child(p, "defRPr"))
            .map(|r| RunProperties::from_xml(r, theme).with_relationships(rels))
            .unwrap_or_default()
            .over(&base)
    })
}
#[cfg(test)]
pub(crate) fn extract_level_run_properties(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> LevelRunProperties {
    extract_level_run_properties_with_rels(tx_body, theme, &HashMap::new())
}

pub(crate) fn extract_level_run_properties_with_rels(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> LevelRunProperties {
    child(tx_body, "lstStyle")
        .map(|n| read_level_run_properties_with_rels(n, theme, rels))
        .unwrap_or_else(|| std::array::from_fn(|_| RunProperties::default()))
}
pub(crate) fn merge_level_run_properties(
    higher: &LevelRunProperties,
    lower: &LevelRunProperties,
) -> LevelRunProperties {
    std::array::from_fn(|i| higher[i].over(&lower[i]))
}
pub(crate) fn has_any_level_run_properties(levels: &LevelRunProperties) -> bool {
    levels.iter().any(|p| !p.is_empty())
}

// Carries the resolved master/layout/placeholder inheritance context (theme,
// rels, inherited font size, default alignment/spacing, level styles) that text
// runs need; these are inheritance inputs, not an arbitrary parameter bag.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_text_body(
    tx_body: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    source_part: &str,
    inherited_font_size: Option<f64>,
    inherited_level_font_sizes: LevelFontSizes,
    inherited_level_colors: LevelColors,
    inherited_level_run_properties: LevelRunProperties,
    inherited_level_indents: LevelIndents,
    inherited_level_bullets: &LevelBullets,
    inherited_bold: Option<bool>,
    inherited_italic: Option<bool>,
    inherited_caps: Option<String>,
    inherited_reflection: Option<Reflection>,
    inherited_anchor: Option<String>,
    inherited_body_pr: Option<InheritedBodyPr>,
    inherited_alignment: Option<String>,
    inherited_level_alignments: &LevelAlignments,
    inherited_ea_ln_brk: Option<bool>,
    inherited_font_algn: Option<String>,
    inherited_spacing: LevelSpacing,
    implicit_mar_l: [i64; 9],
    zip: &mut PptxZip,
) -> TextBody {
    let body_pr = child(tx_body, "bodyPr");
    // Theme `<a:objectDefaults>` (ECMA-376 §20.1.6.7: spDef / lnDef / txDef)
    // are deliberately NOT a fallback here. They are the templates PowerPoint
    // uses for objects newly inserted in its UI; an existing shape that omits
    // a bodyPr attribute resolves it from its placeholder cascade and then the
    // CT_TextBodyProperties schema default. Observed with PowerPoint for Mac
    // PDF export (issue #1618): identical shape XML rendered under a theme
    // whose txDef/spDef set large insets, anchor ctr/b, wrap none, vert270,
    // numCol 2, spcFirstLastPara 1, spAutoFit / normAutofit fontScale 50 %,
    // lstStyle size/colour/face/alignment, spPr fill+line and style refs (and
    // lnDef line + style) matched a control theme with an empty
    // objectDefaults on every case — text boxes (txBox="1"), autoshapes with
    // and without p:style, body placeholders whose layout/master bodyPr omit
    // the attributes, connectors and line shapes. Explicit attributes on the
    // same shapes (fontScale, spcFirstLastPara) did take effect, so the
    // properties were observable. Evidence boundary: an explicit spAutoFit is
    // not re-run on export either, so a txDef spAutoFit could not change the
    // PDF; it is excluded for consistency with every observable property.

    // Shared `<a:bodyPr>` grammar (anchor / wrap / vert / insets / autofit) via
    // ooxml_common::text::parse_body_pr. pptx's placeholder inheritance is
    // pre-baked into the defaults: each field is `inherited?.or(spec default)`,
    // and parse_body_pr then applies the shape's own bodyPr attribute over it,
    // giving the precedence shape attr → inherited → spec. When the shape has
    // no `<a:bodyPr>` at all, the resolved defaults ARE the result.
    //
    // Insets: OOXML defaults lIns=rIns=91440, tIns=bIns=45720 (the shared
    // ooxml_common::text::DEFAULT_INS_* constants, via BodyPrDefaults::spec()).
    // Autofit child (spAutoFit / normAutofit / noAutofit): when absent, the
    // inherited child (with a normAutofit's stored fontScale / lnSpcReduction,
    // ECMA-376 §21.1.2.1.3, 62500 → 0.625) or the schema default (none).
    let spec = ooxml_common::text::BodyPrDefaults::spec();
    let inherited = inherited_body_pr.unwrap_or_default();
    let inherited_fit = inherited.auto_fit.clone();
    let body_pr_defaults = ooxml_common::text::BodyPrDefaults {
        anchor: inherited_anchor.unwrap_or(spec.anchor),
        wrap: inherited.wrap.clone().unwrap_or(spec.wrap),
        vert: inherited.vert.clone().unwrap_or(spec.vert),
        l_ins: inherited.insets[0].unwrap_or(spec.l_ins),
        t_ins: inherited.insets[1].unwrap_or(spec.t_ins),
        r_ins: inherited.insets[2].unwrap_or(spec.r_ins),
        b_ins: inherited.insets[3].unwrap_or(spec.b_ins),
        auto_fit: inherited_fit
            .as_ref()
            .map(|fit| fit.mode.clone())
            .unwrap_or(spec.auto_fit),
    };
    let body = match body_pr {
        Some(n) => ooxml_common::text::parse_body_pr(n, &body_pr_defaults),
        // No <a:bodyPr>: every field resolves to its default.
        None => ooxml_common::text::BodyPr {
            anchor: body_pr_defaults.anchor.clone(),
            wrap: body_pr_defaults.wrap.clone(),
            vert: body_pr_defaults.vert.clone(),
            l_ins: body_pr_defaults.l_ins,
            t_ins: body_pr_defaults.t_ins,
            r_ins: body_pr_defaults.r_ins,
            b_ins: body_pr_defaults.b_ins,
            auto_fit: body_pr_defaults.auto_fit.clone(),
            font_scale: None,
            ln_spc_reduction: None,
        },
    };
    let vertical_anchor = body.anchor;
    let l_ins = body.l_ins;
    let r_ins = body.r_ins;
    let t_ins = body.t_ins;
    let b_ins = body.b_ins;
    let wrap = body.wrap;
    let vert = body.vert;
    let auto_fit = body.auto_fit;
    // parse_body_pr only reports stored scales from the shape's own autofit
    // child. When the shape has none, the inherited child supplies them.
    let own_fit = body_pr.and_then(ooxml_common::text::parse_autofit);
    let (font_scale, ln_spc_reduction) = match (own_fit, inherited_fit) {
        (Some(_), _) => (body.font_scale, body.ln_spc_reduction),
        (None, Some(fit)) => (fit.font_scale, fit.ln_spc_reduction),
        (None, None) => (None, None),
    };
    // ECMA-376 §20.1.10.34: numCol on <a:bodyPr> tells the renderer to
    // distribute paragraphs across N columns within the shape. Default 1.
    // spcCol is the inter-column gutter in EMU (default 0).
    let own = body_pr
        .map(InheritedBodyPr::from_body_pr)
        .unwrap_or_default();
    let num_col = own.num_col.or(inherited.num_col).unwrap_or(1);
    let spc_col = own.spc_col.or(inherited.spc_col).unwrap_or(0);
    // ECMA-376 §21.1.2.1.1: rtlCol on <a:bodyPr> lays out the text body's
    // columns right-to-left. xsd:boolean, so accept "1"/"true". Shape
    // attribute → spec default (false).
    let rtl_col = own.rtl_col.or(inherited.rtl_col).unwrap_or(false);
    // ECMA-376 §21.1.2.1.1 spcFirstLastPara: shape attribute → spec default
    // (false, edge spacing suppressed).
    let spc_first_last_para = own
        .spc_first_last_para
        .or(inherited.spc_first_last_para)
        .unwrap_or(false);
    // ECMA-376 §21.1.2.1.1 anchorCtr ("centered within the bounding box"
    // perpendicular to the anchor), xsd:boolean default false. Cascaded like
    // the other bodyPr attributes above.
    let anchor_ctr = own.anchor_ctr.or(inherited.anchor_ctr).unwrap_or(false);
    // ECMA-376 §21.1.2.1.1 compatLnSpc ("line spacing ... decided in a
    // simplistic manner using the font scene", schema default false). Carried
    // through the placeholder cascade like the other bodyPr attributes; a
    // non-placeholder shape has no inherited value. PowerPoint's reference
    // (Windows-style) PDF export (#1619 controls, both decks' cascade slides):
    // the value authored on the slide, else the layout, else the master
    // placeholder wins (layout 0 over master 1, slide 0 over master 1), and a
    // text box never takes it from the master or layout. The renderer decides
    // what an effective value means; `None` stays distinguishable from `1`.
    let compat_ln_spc = own.compat_ln_spc.or(inherited.compat_ln_spc);

    // ECMA-376 §20.1.9.19 — `<a:bodyPr><a:prstTxWarp prst="…">` selects a WordArt
    // text-warp envelope (ST_TextShapeType). Its `<a:avLst>` carries `<a:gd>`
    // adjust overrides in adj1/adj2/… order (thousandths of a percent). We record
    // the preset name + adjust values; the renderer maps glyphs through the
    // matching envelope from presetTextWarpDefinitions.xml. `prst="textNoShape"`
    // means "no warp", so it is treated as absent.
    //
    // Schema note: CT_TextBodyProperties (dml-main.xsd) is an xsd:sequence whose
    // FIRST child is prstTxWarp — before EG_TextAutofit (spAutoFit/normAutofit),
    // scene3d, EG_Text3D and extLst. Real Office files always emit it in that
    // position, and PowerPoint IGNORES a prstTxWarp placed later in the
    // sequence. This name-based lookup is deliberately position-independent for
    // robustness, but any fixture/generator we author must emit the schema
    // order or PowerPoint itself will render the text un-warped.
    // A placeholder that omits prstTxWarp inherits its layout/master warp; an
    // explicit `textNoShape` ends the cascade with no warp.
    let text_warp = own.text_warp.clone().or(inherited.text_warp).flatten();

    // Own lstStyle > lvl1pPr, then fall back to layout/master inherited values
    let own_lvl1_ppr = child(tx_body, "lstStyle").and_then(|ls| child(ls, "lvl1pPr"));
    let own_def_rpr = own_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
    let own_effects = own_def_rpr.and_then(|rp| child(rp, "effectLst"));
    let default_reflection = match own_effects {
        Some(effect_lst) => parse_reflection(effect_lst),
        None => inherited_reflection,
    };
    let default_font_size = own_def_rpr
        .and_then(|rp| attr_f64(&rp, "sz"))
        .map(|v| v / 100.0)
        .or(inherited_font_size);
    // Effective per-list-level default sizes: this shape's own lstStyle wins per
    // level, else the layout/master inherited per-level sizes. Paragraphs pick
    // their size by `lvl` so nested bullets shrink (ECMA-376 §21.1.2.4).
    let own_level_sizes = extract_level_font_sizes(tx_body);
    let effective_level_sizes = merge_level_sizes(&own_level_sizes, &inherited_level_font_sizes);
    let own_level_colors = extract_level_colors(tx_body, theme);
    let effective_level_colors = merge_level_colors(&own_level_colors, &inherited_level_colors);
    let own_level_run_properties = extract_level_run_properties_with_rels(tx_body, theme, rels);
    let effective_level_run_properties =
        merge_level_run_properties(&own_level_run_properties, &inherited_level_run_properties);
    // Effective per-list-level indents: this shape's own lstStyle wins per
    // axis/level, else the layout/master inherited per-level indents. A paragraph
    // that omits marL/marR/indent picks them by `lvl` from this cascade before
    // falling back to PowerPoint's hardcoded implicit defaults (§21.1.2.4.13).
    let own_level_indents = extract_level_indents(tx_body);
    let effective_level_indents = merge_level_indents(&own_level_indents, &inherited_level_indents);
    // Effective per-level bullets: own lstStyle wins per level, else inherited
    // layout/master. A paragraph with no explicit bullet resolves its marker (and
    // its hanging-indent defaults) from this by `lvl` (ECMA-376 §19.7.10).
    // A slide text body's own lstStyle `<a:buBlip>` resolves against the slide's
    // rels + actual source-part directory (ECMA-376 §21.1.2.4.2 and Part 2
    // §6.5.2.3), the same base as the containing shape's picture fills.
    let mut resolve_slide_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path(source_part, target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the slide picture-fill resolvers. `index_for_name`
        // reads the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let own_level_bullets = extract_level_bullets(tx_body, theme, &mut resolve_slide_blip);
    let effective_level_bullets = merge_level_bullets(&own_level_bullets, inherited_level_bullets);
    let default_bold = own_def_rpr
        .and_then(|rp| attr(&rp, "b"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_bold);
    let default_italic = own_def_rpr
        .and_then(|rp| attr(&rp, "i"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_italic);
    // Per-level alignment: the own lstStyle level, else the inherited level.
    // A level neither sets falls back to the body default below.
    let own_level_alignments = child(tx_body, "lstStyle")
        .map(read_level_alignments)
        .unwrap_or_default();
    let effective_level_alignments: LevelAlignments = std::array::from_fn(|lvl| {
        own_level_alignments[lvl]
            .clone()
            .or_else(|| inherited_level_alignments[lvl].clone())
    });
    // Own lstStyle > lvl1pPr > algn overrides inherited alignment
    let body_default_alignment = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "algn"))
        .map(|a| a.to_string())
        .or(inherited_alignment);

    // Own lstStyle > lvl1pPr > eaLnBrk overrides inherited (ECMA-376 §21.1.2.2.7)
    let body_default_ea_ln_brk = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(inherited_ea_ln_brk);

    // Own lstStyle > lvl1pPr > fontAlgn overrides inherited (ECMA-376
    // §21.1.2.2.7), mirroring eaLnBrk. A paragraph's own pPr@fontAlgn wins
    // (resolved below, after parse_paragraph read it).
    let body_default_font_algn = own_lvl1_ppr
        .and_then(|lp| attr(&lp, "fontAlgn"))
        .map(|v| v.to_string())
        .or(inherited_font_algn);

    // Own lstStyle levels over the inherited levels, per level and property.
    let own_spacing = child(tx_body, "lstStyle")
        .map(LevelSpacing::read)
        .unwrap_or_default();
    let effective_spacing = own_spacing.or(&inherited_spacing);

    let mut paragraphs: Vec<Paragraph> = children_vec(tx_body, "p")
        .into_iter()
        .map(|p| {
            parse_paragraph(
                p,
                theme,
                rels,
                source_part,
                body_default_alignment.as_deref(),
                &effective_level_alignments,
                body_default_ea_ln_brk,
                &effective_spacing,
                default_reflection.as_ref(),
                &effective_level_sizes,
                &effective_level_run_properties,
                &effective_level_indents,
                &effective_level_bullets,
                implicit_mar_l,
                zip,
            )
        })
        .collect();
    for para in &mut paragraphs {
        para.font_algn =
            effective_font_algn(para.font_algn.take(), body_default_font_algn.as_deref());
    }

    // A paragraph's own pPr > defRPr remains the most specific colour. When
    // absent, inherit the defRPr fill from the matching list level rather than
    // the text body's lvl1 default. `pPr@lvl="1"` selects lvl2pPr.
    for paragraph in &mut paragraphs {
        if paragraph.def_color.is_none() {
            paragraph.def_color = effective_level_colors
                .get(paragraph.lvl as usize)
                .and_then(Clone::clone);
        }
    }

    // ECMA-376 §21.1.2.3.9, ST_TextCapsType §20.1.10.64: a run inherits
    // cap="all"/"small" from the shape's own lstStyle defRPr, else from the
    // layout/master placeholder
    // style (e.g. a template's titleStyle cap="all" upper-cases the title even
    // though the run's text is stored mixed-case). Run-level rPr/paragraph
    // defRPr already won via parse_run; fill the remainder here.
    let body_caps = own_def_rpr
        .and_then(|rp| attr(&rp, "cap"))
        .filter(|v| v == "all" || v == "small")
        .or(inherited_caps);
    if let Some(bc) = body_caps {
        for para in &mut paragraphs {
            for run in &mut para.runs {
                if let TextRun::Text(t) = run {
                    if t.caps.is_none() {
                        t.caps = Some(bc.clone());
                    }
                }
            }
        }
    }

    TextBody {
        vertical_anchor,
        paragraphs,
        default_font_size,
        default_bold,
        default_italic,
        l_ins,
        r_ins,
        t_ins,
        b_ins,
        wrap,
        vert,
        auto_fit,
        font_scale,
        ln_spc_reduction,
        num_col,
        spc_col,
        rtl_col,
        spc_first_last_para,
        anchor_ctr,
        compat_ln_spc,
        text_warp,
    }
}

/// Walk `node` for OMML math and push a `TextRun::Math` for each equation,
/// descending PowerPoint's `mc:AlternateContent` / `mc:Choice` / `a14:m`
/// wrappers (ECMA-376 §22.1; the a14 markup is from the 2010 drawing ext).
/// Find the font size (pt) of an equation from the first run property within it
/// that carries `sz`. PowerPoint puts the size on the math run's `a:rPr` (or
/// `m:rPr`) rather than the paragraph defRPr, so inline math matches the
/// surrounding text size. `sz` is in hundredths of a point (ECMA-376 §21.1.2.3.9).
pub(crate) fn math_run_size(om: roxmltree::Node<'_, '_>) -> Option<f64> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| attr_f64(&rpr, "sz"))
        .map(|v| v / 100.0)
}

/// Equation colour: the first run-property solidFill within the equation
/// (PowerPoint puts the colour on the math run's `a:rPr`, like the size), so
/// inline math follows the surrounding text colour (e.g. a purple title).
pub(crate) fn math_run_color(
    om: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> Option<String> {
    om.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rPr")
        .find_map(|rpr| child(rpr, "solidFill").and_then(|sf| parse_color_node(sf, theme)))
}

pub(crate) fn push_math_runs(
    node: roxmltree::Node<'_, '_>,
    font_size: Option<f64>,
    theme: &HashMap<String, String>,
    runs: &mut Vec<TextRun>,
) {
    match node.tag_name().name() {
        "oMath" => {
            let nodes = parse_omath_nodes(node);
            if !nodes.is_empty() {
                runs.push(TextRun::Math {
                    nodes,
                    display: false,
                    font_size: math_run_size(node).or(font_size),
                    color: math_run_color(node, theme),
                });
            }
        }
        "oMathPara" => {
            for om in node
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "oMath")
            {
                let nodes = parse_omath_nodes(om);
                if !nodes.is_empty() {
                    runs.push(TextRun::Math {
                        nodes,
                        display: true,
                        font_size: math_run_size(om).or(font_size),
                        color: math_run_color(om, theme),
                    });
                }
            }
        }
        "AlternateContent" => {
            // §9.3 — take the understood Choice's live equation, else the Fallback
            // (which for math is a raster this walker ignores).
            if let Some(sel) = ooxml_common::mce::select_alternate_content(
                node,
                &crate::shape::pptx_understands_ns,
            ) {
                for c in sel.children().filter(|n| n.is_element()) {
                    push_math_runs(c, font_size, theme, runs);
                }
            }
        }
        // a14:m wrapper (local name "m") holds an m:oMathPara.
        "m" => {
            for c in node.children().filter(|n| n.is_element()) {
                push_math_runs(c, font_size, theme, runs);
            }
        }
        _ => {}
    }
}

// Same inherited paragraph/run context as parse_text_body, scoped to one <a:p>.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_paragraph(
    p_node: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
    source_part: &str,
    body_default_alignment: Option<&str>,
    level_alignments: &LevelAlignments,
    body_default_ea_ln_brk: Option<bool>,
    level_spacing: &LevelSpacing,
    body_default_reflection: Option<&Reflection>,
    level_font_sizes: &LevelFontSizes,
    level_run_properties: &LevelRunProperties,
    level_indents: &LevelIndents,
    level_bullets: &LevelBullets,
    implicit_mar_l: [i64; 9],
    zip: &mut PptxZip,
) -> Paragraph {
    let p_pr = child(p_node, "pPr");

    // ECMA-376 §21.1.2.2.7 `<a:pPr rtl>` — right-to-left text flow. When set
    // and the paragraph has no explicit `algn`, the implicit default flips
    // from "l" to "r" (matches PowerPoint's behaviour for Arabic / Hebrew
    // slides where users typically don't author an explicit alignment).
    let rtl = p_pr
        .and_then(|n| attr(&n, "rtl"))
        .map(|v| v == "1" || v == "true")
        .unwrap_or(false);

    // ECMA-376 §21.1.2.2.7 `<a:pPr eaLnBrk>` (xsd:boolean). Paragraph's own
    // value → body/list-style → layout/master default → spec default (true).
    // Same inheritance shape as `alignment` above.
    let ea_ln_brk = p_pr
        .and_then(|n| attr(&n, "eaLnBrk"))
        .map(|v| v == "1" || v == "true")
        .or(body_default_ea_ln_brk)
        .unwrap_or(true);

    // `<a:pPr fontAlgn>` as authored on this paragraph; parse_text_body
    // completes the cascade with `effective_font_algn`.
    let font_algn = p_pr
        .and_then(|n| attr(&n, "fontAlgn"))
        .map(|v| v.to_string());

    // Paragraph's own algn → body/layout/master default → "r" if rtl, else "l"
    let lvl: u32 = p_pr
        .and_then(|n| attr(&n, "lvl"))
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let alignment = p_pr
        .and_then(|n| attr(&n, "algn"))
        .or_else(|| level_alignments[(lvl as usize).min(8)].clone())
        .or_else(|| body_default_alignment.map(|a| a.to_string()))
        .unwrap_or_else(|| if rtl { "r".into() } else { "l".into() });

    // Effective bullet: the paragraph's own bullet groups
    // (`<a:buClr>`/`<a:buSz…>`/`<a:buFont>` + `<a:buChar>`/`<a:buAutoNum>`/
    // `<a:buBlip>`/`<a:buNone>`) merged PER-GROUP over the inherited per-level
    // bullet for this placeholder (ECMA-376 §19.7.10, §21.1.2.4). Each group
    // resolves independently, so a paragraph that sets only `buClr` still
    // inherits the level's marker (and vice versa). A paragraph's own `<a:buBlip>`
    // embed resolves against the containing part's rels + actual source
    // directory, the same base as its picture fills (§21.1.2.4.2; Part 2
    // §6.5.2.3).
    let mut resolve_para_blip = |rid: &str| -> Option<String> {
        let target = rels.get(rid)?;
        let path = resolve_path(source_part, target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // buBlip marker falls through (inherit), mirroring the slide picture-fill
        // resolvers. `index_for_name` reads the central directory only (no
        // inflate), unlike the former `read_zip_bytes` which decompressed the
        // entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };
    let own_bullet = parse_bullet_props(p_pr, theme, &mut resolve_para_blip);
    let inherited_bullet = level_bullets.get(lvl as usize).cloned().unwrap_or_default();
    let bullet_props = BulletProps::merge(&own_bullet, &inherited_bullet);
    // A paragraph is a list item (and gets a hanging indent) when its RESOLVED
    // marker is a char/number/picture — whether declared explicitly or inherited.
    // An inherited bullet without an inherited marL/indent reuses PowerPoint's
    // implicit list metrics, the same defaults explicit bullets already use.
    let has_bullet = matches!(
        bullet_props.marker,
        Some(BuMarker::Char(_)) | Some(BuMarker::AutoNum { .. }) | Some(BuMarker::Blip { .. })
    );
    let bullet = bullet_props.resolve();

    // marL / marR / indent resolve per axis: direct `<a:pPr>` attribute wins,
    // else the authored list-style level cascade (`level_indents`, from the
    // shape/layout/master lstStyle per ECMA-376 §21.1.2.4.13), else the
    // implicit defaults:
    //   Bullet paragraphs:  marL = (lvl+1)*342900, indent = -342900 (hanging)
    //   Plain paragraphs:   the caller's `implicit_mar_l` for the level. A
    //                       placeholder passes 0, a text box its
    //                       defaultTextStyle level (see `DefaultTextLevels`).
    let level_indent = level_indents.get(lvl as usize).copied().unwrap_or_default();
    let mar_l = p_pr
        .and_then(|n| attr_i64(&n, "marL"))
        .or(level_indent.mar_l)
        .unwrap_or_else(|| {
            if has_bullet {
                (lvl as i64 + 1) * 342900
            } else {
                implicit_mar_l[(lvl as usize).min(8)]
            }
        });
    let mar_r = p_pr
        .and_then(|n| attr_i64(&n, "marR"))
        .or(level_indent.mar_r)
        .unwrap_or(0);
    let indent = p_pr
        .and_then(|n| attr_i64(&n, "indent"))
        .or(level_indent.indent)
        .unwrap_or(if has_bullet { -342900 } else { 0 });

    // The nearest `<a:spcBef>`/`<a:spcAft>` wins as a whole: a percentage
    // replaces an inherited point value and vice versa (xsd:choice).
    let (space_before, space_before_pct) = ParagraphSpacing::split(
        p_pr.and_then(|n| paragraph_spacing(n, "spcBef"))
            .or(level_spacing.before[lvl.min(8) as usize]),
    );
    let (space_after, space_after_pct) = ParagraphSpacing::split(
        p_pr.and_then(|n| paragraph_spacing(n, "spcAft"))
            .or(level_spacing.after[lvl.min(8) as usize]),
    );

    let space_line = p_pr
        .and_then(|n| child(n, "lnSpc"))
        .and_then(parse_lnspc)
        .or_else(|| level_spacing.line[lvl.min(8) as usize].clone());

    // Tab stops from pPr > tabLst
    let tab_stops: Vec<TabStop> = p_pr
        .and_then(|n| child(n, "tabLst"))
        .map(|tab_lst| {
            tab_lst
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "tab")
                .filter_map(|tab| {
                    let pos = attr_i64(&tab, "pos")?;
                    let algn = attr(&tab, "algn").unwrap_or_else(|| "l".into());
                    Some(TabStop { pos, algn })
                })
                .collect()
        })
        .unwrap_or_default();

    // §21.1.2.2.7 defTabSz — the paragraph's default tab interval (EMU). When a
    // `\t` has no reachable explicit stop it snaps to this grid (issue #1006).
    // Read only from the paragraph's own pPr; a value inherited through
    // lstStyle/lvlNpPr/layout/master is not resolved here. Acceptable because the
    // renderer falls back to PowerPoint's universal 1-inch default (914400 EMU) —
    // the value every real deck's defaultTextStyle carries.
    let def_tab_sz = p_pr
        .and_then(|n| attr_i64(&n, "defTabSz"))
        .filter(|&v| v > 0);

    // ECMA-376 §21.1.2.2.7 / §21.1.2.4: pPr/defRPr overlays the
    // corresponding lstStyle level property by property. PowerPoint PDF
    // confirms a paragraph with only b="1" still inherits that level's
    // patterned fill for both a:r and a:fld.
    let paragraph_props = p_pr
        .and_then(|n| child(n, "defRPr"))
        .map(|n| RunProperties::from_xml(n, theme).with_relationships(rels))
        .unwrap_or_default();
    let mut defaults = paragraph_props.over(&level_run_properties[lvl.min(8) as usize]);
    if !defaults.effects_authored && defaults.reflection.is_none() {
        defaults.reflection = body_default_reflection.cloned();
    }
    let def_font_size = defaults.font_size;
    let def_color = defaults.color.clone();
    let def_bold = defaults.bold;
    let def_italic = defaults.italic;
    // The list-level face already ends the placeholder / defaultTextStyle
    // chain (shape.rs); levels never borrow the level-1 face (#1620).
    let def_font_family =
        finalize_latin_face(defaults.font_family.as_deref(), theme, defaults.language());

    let mut runs = Vec::new();
    for node in p_node.children().filter(|n| n.is_element()) {
        match node.tag_name().name() {
            "r" => {
                if let Some(run) = parse_run_with_defaults(node, &defaults, theme, rels) {
                    runs.push(TextRun::Text(run));
                }
            }
            "br" => {
                let br_pr = child(node, "rPr");
                let own = br_pr
                    .map(|r| RunProperties::from_xml(r, theme).with_relationships(rels))
                    .unwrap_or_default();
                let effective = own.over(&defaults);
                runs.push(TextRun::Break {
                    // Only a size/face/weight AUTHORED on the break adds a
                    // line-metric segment.  A language-only rPr inherits the
                    // paragraph's style for editing but must not inflate the
                    // preceding line to the master default size (observed in
                    // PowerPoint-exported Japanese and title controls).
                    font_size: own.font_size,
                    font_family: finalize_latin_face(
                        own.font_family.as_deref(),
                        theme,
                        effective.language(),
                    ),
                    bold: own.bold,
                    italic: own.italic,
                    character_attributes: br_pr.map(|_| effective.attributes).unwrap_or_default(),
                    character_child_attributes: br_pr
                        .map(|_| effective.child_attributes)
                        .unwrap_or_default(),
                });
            }
            // OMML equations (ECMA-376 §22.1). `def_font_size` here is the
            // paragraph's defRPr size (pre-level-fallback); the renderer applies
            // its own inheritance when this is None. PowerPoint stores inline
            // math as a bare `a14:m` (local name "m") directly under `a:p`, and
            // also inside `mc:AlternateContent`; both reach push_math_runs.
            "oMath" | "oMathPara" | "AlternateContent" | "m" => {
                push_math_runs(node, def_font_size, theme, &mut runs);
            }
            // ECMA-376 §21.1.2.2.7: a:fld has the same rPr/t content as a:r.
            // Resolve its formatting through the run cascade so defRPr and
            // list-level defaults apply to every property, not only text fill.
            "fld" => {
                if let Some(mut run) = parse_run_with_defaults(node, &defaults, theme, rels) {
                    if attr(&node, "type").as_deref() == Some("slidenum") {
                        run.field_type = Some("slidenum".to_string());
                    }
                    runs.push(TextRun::Text(run));
                }
            }
            _ => {}
        }
    }

    // For paragraphs with no visible text content, use endParaRPr sz to set line height.
    // This ensures empty spacer paragraphs have the correct height (e.g. between sections).
    let end_rpr = child(p_node, "endParaRPr");
    let end_run_properties = end_rpr.map(|node| {
        Box::new(resolve_run_properties(
            String::new(),
            RunProperties::from_xml(node, theme).with_relationships(rels),
            &defaults,
            theme,
        ))
    });
    let has_text = runs
        .iter()
        .any(|r| matches!(r, TextRun::Text(t) if !t.text.is_empty()));
    // endParaRPr is the formatting of the insertion position after the last
    // character (§21.1.2.2.2), not another default for existing runs.  For an
    // empty paragraph that position is its only line, so its authored size
    // wins over inherited defaults.  This also keeps empty spacer paragraphs
    // at their Office height when a master supplies a different size.
    let def_font_size = (if !has_text {
        end_rpr.and_then(|n| attr_f64(&n, "sz")).map(|v| v / 100.0)
    } else {
        None
    })
    .or(def_font_size)
    // Inherited per-list-level default size, indexed by this paragraph's
    // level (ECMA-376 §21.1.2.4): a 2nd-level bullet uses lvl3pPr's smaller
    // defRPr sz, not the level-1 size. The renderer applies `def_font_size`
    // to runs that carry no explicit `sz`.
    .or_else(|| level_font_sizes.get(lvl as usize).copied().flatten());

    Paragraph {
        alignment,
        mar_l,
        mar_r,
        indent,
        space_before,
        space_after,
        space_before_pct,
        space_after_pct,
        space_line,
        lvl,
        bullet,
        def_font_size,
        def_color,
        def_bold,
        def_italic,
        def_font_family,
        tab_stops,
        def_tab_sz,
        rtl,
        ea_ln_brk,
        font_algn,
        runs,
        end_run_properties,
    }
}

/// The effective `fontAlgn` (ST_TextFontAlignType, ECMA-376 §20.1.10.62):
/// the paragraph's own value, else the body/layout/master default. Only the
/// values that change PowerPoint's layout are kept: an omitted value, `auto`
/// and `base` render identically (#1619 controls in both line models), and an
/// unknown token is ignored like an omitted one.
pub(crate) fn effective_font_algn(own: Option<String>, inherited: Option<&str>) -> Option<String> {
    let valid = |v: &str| matches!(v, "auto" | "t" | "ctr" | "base" | "b");
    let resolved = own
        .filter(|v| valid(v))
        .or_else(|| inherited.filter(|v| valid(v)).map(str::to_string))?;
    matches!(resolved.as_str(), "t" | "ctr" | "b").then_some(resolved)
}

/// Parse the marker choice group (ECMA-376 §21.1.2.4 EG_TextBullet) from a pPr /
/// lvlNpPr node. The four members are an `xsd:choice`; PowerPoint files carry at
/// most one, but we keep the historical precedence `buNone > buBlip > buChar >
/// buAutoNum` for robustness against malformed inputs.
///
/// `resolve_blip` maps a `<a:buBlip><a:blip r:embed>` rId to the bullet image's
/// embedded **zip path** (§21.1.2.4.2), using the rels + part directory of
/// whichever tier this node belongs to (slide paragraph / txBody lstStyle /
/// layout / master), mirroring how `parse_blip_fill` resolves image fills. A
/// `buBlip` whose embed can't be resolved (dangling rId) yields `None` (no
/// marker) so a lower style tier can still supply one.
fn parse_bullet_marker<F: FnMut(&str) -> Option<String>>(
    p_pr: roxmltree::Node<'_, '_>,
    resolve_blip: &mut F,
) -> Option<BuMarker> {
    // Explicit "no bullet"
    if child(p_pr, "buNone").is_some() {
        return Some(BuMarker::None);
    }

    // Picture bullet (buBlip) — only a resolvable embed emits a marker.
    if let Some(bu_blip) = child(p_pr, "buBlip") {
        if let Some(image_path) = child(bu_blip, "blip")
            .and_then(|b| attr_r(&b, "embed"))
            .and_then(|rid| resolve_blip(&rid))
        {
            let mime_type = mime_from_ext(&image_path).to_owned();
            return Some(BuMarker::Blip {
                image_path,
                mime_type,
            });
        }
        // Dangling embed: fall through so a lower tier's marker can supply one.
    }

    // Character bullet
    if let Some(bu_char) = child(p_pr, "buChar") {
        // CT_TextCharBullet's attribute is schema-typed as a string. Flattened
        // SmartArt caches can nevertheless repeat one marker (for example
        // `••`) even though PowerPoint paints it once. Collapse only that
        // duplicate-marker form; preserve genuinely multi-character values.
        let ch = attr(&bu_char, "char")
            .map(|value| {
                let mut chars = value.chars();
                let first = chars.next();
                match first {
                    Some(marker) if chars.clone().count() > 0 && chars.all(|ch| ch == marker) => {
                        marker.to_string()
                    }
                    _ => value,
                }
            })
            .unwrap_or_else(|| "\u{2022}".into()); // •
        return Some(BuMarker::Char(ch));
    }

    // Auto-numbered bullet
    if let Some(bu_auto) = child(p_pr, "buAutoNum") {
        let num_type = attr(&bu_auto, "type").unwrap_or_else(|| "arabicPeriod".into());
        let start_at = attr(&bu_auto, "startAt").and_then(|v| v.parse().ok());
        return Some(BuMarker::AutoNum { num_type, start_at });
    }

    None
}

/// Parse a `<a:buSzPct val>` value into a percentage of the text size (100.0 =
/// 100%). ECMA-376's `ST_TextBulletSizePercent` has two lexical forms: the
/// Transitional integer in thousandths of a percent (`"100000"` = 100%, what
/// PowerPoint writes) and the Strict percentage string (`"111%"`, as in the
/// spec's own example, §21.1.2.4.9). Accept both; a trailing `%` selects the
/// direct-percentage reading.
pub(crate) fn parse_bu_sz_pct(val: &str) -> Option<f64> {
    let v = val.trim();
    match v.strip_suffix('%') {
        Some(pct) => pct.trim().parse::<f64>().ok(),
        None => v.parse::<f64>().ok().map(|n| n / 1000.0),
    }
}

/// Parse a `<a:buSzPts val>` value into an absolute size in points. ECMA-376's
/// attribute type is `ST_TextFontSize` (§21.1.2.4.10 CT_TextBulletSizePoint):
/// an integer in hundredths of a point (`"1800"` = 18pt), exactly like the run
/// `<a:rPr sz>` the parser already divides by 100 elsewhere.
pub(crate) fn parse_bu_sz_pts(val: &str) -> Option<f64> {
    val.trim().parse::<f64>().ok().map(|n| n / 100.0)
}

/// Parse a paragraph's four bullet choice groups (ECMA-376 §21.1.2.4) from a
/// `<a:pPr>` / `<a:lvlNpPr>` node into a [`BulletProps`]. Each group is read
/// independently so the cascade can merge them per-property:
/// - colour (EG_TextBulletColor): `<a:buClrTx>` (follow text) or `<a:buClr>`;
/// - size (EG_TextBulletSize): `<a:buSzTx>` (follow text), `<a:buSzPct>` (percent
///   of text) or `<a:buSzPts>` (absolute points, §21.1.2.4.10) — see [`BuSize`];
/// - typeface (EG_TextBulletTypeface): `<a:buFontTx>` (follow text) or `<a:buFont>`;
/// - marker (EG_TextBullet): see [`parse_bullet_marker`].
///
/// A `None` node (no `<a:pPr>`) yields all-inherit. Each group defaults to `None`
/// (inherit) when its elements are absent, so a lower style tier supplies it.
pub(crate) fn parse_bullet_props<F: FnMut(&str) -> Option<String>>(
    p_pr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> BulletProps {
    let p_pr = match p_pr {
        Some(n) => n,
        None => return BulletProps::default(),
    };

    // Colour group (EG_TextBulletColor): buClrTx (§21.1.2.4.5) breaks inheritance
    // and follows the run colour; buClr (§21.1.2.4.4) is explicit.
    let color = if child(p_pr, "buClrTx").is_some() {
        Some(BuColor::FollowText)
    } else {
        child(p_pr, "buClr")
            .and_then(|n| parse_color_node(n, theme))
            .map(BuColor::Color)
    };

    // Size group (EG_TextBulletSize xsd:choice): buSzTx (§21.1.2.4.11) follow-text;
    // buSzPct (§21.1.2.4.9) percent-of-text; buSzPts (§21.1.2.4.10) absolute points.
    // At most one is present; check them in schema order so a well-formed level
    // reads deterministically (a stray second element is ignored).
    let size = if child(p_pr, "buSzTx").is_some() {
        Some(BuSize::FollowText)
    } else if let Some(pct) = child(p_pr, "buSzPct")
        .and_then(|n| attr(&n, "val"))
        .and_then(|v| parse_bu_sz_pct(&v))
    {
        Some(BuSize::Pct(pct))
    } else {
        child(p_pr, "buSzPts")
            .and_then(|n| attr(&n, "val"))
            .and_then(|v| parse_bu_sz_pts(&v))
            .map(BuSize::Pts)
    };

    // Typeface group (EG_TextBulletTypeface): buFontTx (§21.1.2.4.7) follow-text;
    // buFont (§21.1.2.4.6) explicit typeface.
    let font = if child(p_pr, "buFontTx").is_some() {
        Some(BuFont::FollowText)
    } else {
        child(p_pr, "buFont")
            .and_then(|n| attr(&n, "typeface"))
            .map(|tf| BuFont::Font(resolve_theme_typeface(&tf, theme)))
    };

    let marker = parse_bullet_marker(p_pr, resolve_blip);

    BulletProps {
        marker,
        color,
        font,
        size,
    }
}

/// Test helper: parse a single-tier `<a:pPr>` bullet and collapse it to the
/// serialized [`Bullet`]. Production code cascades [`BulletProps`] across tiers
/// and only collapses at the leaf ([`BulletProps::resolve`]).
#[cfg(test)]
pub(crate) fn parse_bullet<F: FnMut(&str) -> Option<String>>(
    p_pr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    resolve_blip: &mut F,
) -> Bullet {
    parse_bullet_props(p_pr, theme, resolve_blip).resolve()
}

/// Parse a standalone text run without a shape/master-level reflection input.
///
/// Most call sites (and the parser's focused tests) only need run-local and
/// paragraph `defRPr` inheritance. Text-body parsing uses the private extended
/// form below because placeholder title styles can additionally contribute a
/// reflection through the master/layout cascade.
#[cfg(test)]
pub(crate) fn parse_run(
    r_node: roxmltree::Node<'_, '_>,
    def_rpr: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> Option<TextRunData> {
    let defaults = def_rpr
        .map(|n| RunProperties::from_xml(n, theme).with_relationships(rels))
        .unwrap_or_default();
    parse_run_with_defaults(r_node, &defaults, theme, rels)
}

fn parse_run_with_defaults(
    r_node: roxmltree::Node<'_, '_>,
    defaults: &RunProperties,
    theme: &HashMap<String, String>,
    rels: &HashMap<String, String>,
) -> Option<TextRunData> {
    let t_node = child(r_node, "t");
    if t_node.is_none() && r_node.tag_name().name() != "fld" {
        return None;
    }
    // CT_TextField permits an absent a:t; a slide-number field still needs
    // to reach the renderer so it can substitute the current slide number.
    let text = t_node.and_then(|n| n.text()).unwrap_or("").to_owned();
    let r_pr = child(r_node, "rPr");
    let authored = r_pr
        .map(|n| RunProperties::from_xml(n, theme).with_relationships(rels))
        .unwrap_or_default();
    Some(resolve_run_properties(text, authored, defaults, theme))
}

fn inherited_hyperlink_target(
    relationship: &Option<Option<InheritedRelationship>>,
    attributes: Option<&PropertyAttributes>,
) -> Option<String> {
    let rel = relationship.as_ref()?.as_ref()?;
    if rel.target.is_empty() {
        return None;
    }
    // An inherited @action can come from a different level than its r:id.
    // Resolve a slide jump against the ID owner's OPC part only after the
    // complete cascade. A relative external URI ending in `.xml` stays raw.
    if attributes
        .and_then(|a| a.get("action"))
        .is_some_and(|action| action == "ppaction://hlinksldjump")
        && rel.target.ends_with(".xml")
        && !rel.target.contains("://")
    {
        if let Some(source_part) = &rel.source_part {
            return Some(resolve_path(source_part, &rel.target));
        }
    }
    Some(rel.target.clone())
}

/// Resolve the same character-property chain for an a:r, a:fld, or the
/// trailing endParaRPr insertion state.  The latter is stored separately and
/// never painted over an existing run (§21.1.2.2.2).
fn resolve_run_properties(
    text: String,
    authored: RunProperties,
    defaults: &RunProperties,
    theme: &HashMap<String, String>,
) -> TextRunData {
    let props = authored.over(defaults);
    let lang = props.language();
    let alt_lang = props.alt_language();
    let underline = props.underline.as_deref().is_some_and(|v| v != "none");
    let underline_style = props
        .underline
        .clone()
        .filter(|v| v != "none" && v != "sng");
    let underline_fill = props.underline_fill.clone().flatten();
    let underline_color = match &underline_fill {
        Some(Fill::Solid { color }) => Some(color.clone()),
        _ => None,
    };
    let strikethrough = matches!(props.strike.as_deref(), Some("sngStrike" | "dblStrike"));
    let strike_double = props.strike.as_deref() == Some("dblStrike");
    // An explicit cap="none" blocks an inherited all/small value.  The
    // renderer treats "none" as a no-op while retaining its presence.
    let caps = props.caps.clone();
    let letter_spacing = props.letter_spacing.filter(|v| v.abs() > f64::EPSILON);
    let pattern_fill = match &props.fill {
        Some(fill @ Fill::Pattern { .. }) => Some(fill.clone()),
        _ => None,
    };
    let no_fill = matches!(props.fill, Some(Fill::None));
    let color = props.color.clone();
    // Theme tokens resolve per run language (issue #1627). An empty ea/cs
    // slot stays None: the renderer then applies Office's application
    // default for the script, never the Latin face.
    let font_family = finalize_latin_face(props.font_family.as_deref(), theme, lang);
    let font_family_ea = props
        .font_family_ea
        .as_deref()
        .and_then(|face| resolve_slot_face(face, FontSlot::EastAsian, theme, lang, alt_lang));
    let font_family_cs = props
        .font_family_cs
        .as_deref()
        .and_then(|face| resolve_slot_face(face, FontSlot::ComplexScript, theme, lang, alt_lang));
    let run_lang = lang.map(str::to_owned);
    let run_alt_lang = alt_lang.map(str::to_owned);
    let font_family_sym = props.font_family_sym.clone().filter(|v| !v.is_empty());
    let baseline = props.baseline.filter(|v| *v != 0);

    // a:hlinkClick — hyperlink. Its r:id was resolved against the owning
    // master, layout, or slide part before character-property inheritance;
    // the renderer therefore needs no rels table. ECMA-376 §21.1.2.3.5:
    // optional @action holds a "ppaction://..." verb (e.g. hlinksldjump) that
    // marks the link as an INTERNAL navigation; carry it through so the TS side
    // can distinguish a slide jump from an external URL. For a slide jump the
    // rel is TargetMode=Internal, so `hyperlink` is the internal slide part.
    let hlink_click = props.child_attributes.get("hlinkClick");
    let hyperlink = inherited_hyperlink_target(&props.hlink_click_target, hlink_click);
    let hyperlink_action = hlink_click
        .and_then(|h| h.get("action").cloned())
        .filter(|s| !s.is_empty());
    let hlink_mouse_over = props.child_attributes.get("hlinkMouseOver");
    let hyperlink_mouse_over =
        inherited_hyperlink_target(&props.hlink_mouse_over_target, hlink_mouse_over);
    let hyperlink_mouse_over_action = hlink_mouse_over
        .and_then(|h| h.get("action").cloned())
        .filter(|s| !s.is_empty());
    // Office's hyperlink-colour extension is written when an authored text
    // fill is reapplied after creating the link. Without hlinkClr="tx", the
    // hyperlink theme colour wins. The extension is scoped to run hyperlinks.
    let hyperlink_uses_text_fill = props.hyperlink_uses_text_fill.unwrap_or(false);
    // PowerPoint's hyperlink theme colour wins when the run has no authored
    // fill of its own, even if a list-style defRPr supplies a solid colour.
    // Preserve run-local solid colours, and hlinkClr="tx" explicitly asks to
    // keep the inherited text paint. Observed with both list-style links and
    // directly coloured links in PowerPoint PDF output.
    let color = if hyperlink.is_some() && !hyperlink_uses_text_fill && !authored.fill_authored {
        None
    } else {
        color
    };

    TextRunData {
        text,
        bold: props.bold,
        italic: props.italic,
        underline,
        underline_style,
        underline_color,
        underline_fill,
        underline_line: props.underline_line.clone().flatten(),
        underline_line_no_fill: props.underline_line.as_ref().is_some_and(Option::is_none)
            && !props.underline_line_follow_text.unwrap_or(false),
        strikethrough,
        strike_double,
        font_size: props.font_size,
        color,
        pattern_fill,
        glyph_fill: props.fill.clone(),
        no_fill,
        font_family,
        font_family_ea,
        font_family_cs,
        font_family_sym,
        lang: run_lang,
        alt_lang: run_alt_lang,
        baseline,
        caps,
        letter_spacing,
        field_type: None,
        hyperlink,
        hyperlink_uses_text_fill,
        hyperlink_action,
        hyperlink_mouse_over,
        hyperlink_mouse_over_action,
        shadow: props.shadow,
        reflection: props.reflection,
        outline: props.outline.flatten(),
        highlight: props.highlight,
        character_attributes: props.attributes,
        character_child_attributes: props.child_attributes,
    }
}

#[cfg(test)]
mod relationship_owner_tests {
    use super::*;
    use crate::master::parse_master_level_run_properties;

    #[test]
    fn inherited_hyperlinks_resolve_against_the_part_that_authored_the_id() {
        // OPC relationship IDs have part-local scope.  All three parts use
        // rId7, and an attribute-only child must not change its owner's ID.
        let master = roxmltree::Document::parse(
            r#"
          <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr>
              <a:hlinkClick r:id="rId7" tooltip="master"/>
              <a:hlinkMouseOver r:id="rId7"/>
            </a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles>
          </p:sldMaster>"#,
        )
        .unwrap();
        let layout = roxmltree::Document::parse(
            r#"
          <a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <a:lstStyle><a:lvl1pPr><a:defRPr>
              <a:hlinkClick r:id="rId7"/>
              <a:hlinkMouseOver r:id="rId7"/>
            </a:defRPr></a:lvl1pPr></a:lstStyle>
          </a:txBody>"#,
        )
        .unwrap();
        let master_rels = HashMap::from([("rId7".into(), "https://master.test/".into())]);
        let layout_rels = HashMap::from([("rId7".into(), "https://layout.test/".into())]);
        let slide_rels = HashMap::from([("rId7".into(), "https://slide.test/".into())]);
        let theme = HashMap::new();
        let master_levels = parse_master_level_run_properties(
            master.root_element(),
            &theme,
            &master_rels,
            "ppt/slideMasters/slideMaster1.xml",
        );
        let master_default = &master_levels.placeholders["body"][0];
        let layout_levels =
            extract_level_run_properties_with_rels(layout.root_element(), &theme, &layout_rels);
        let layout_default = layout_levels[0].over(master_default);

        for (xml, defaults, click, hover) in [
            ("<r><rPr/><t>master</t></r>", master_default,
             "https://master.test/", "https://master.test/"),
            ("<r><rPr><hlinkClick tooltip=\"near\"/></rPr><t>layout</t></r>", &layout_default,
             "https://layout.test/", "https://layout.test/"),
            ("<r><rPr><hlinkClick r:id=\"rId7\"/><hlinkMouseOver r:id=\"rId7\"/></rPr><t>slide</t></r>", &layout_default,
             "https://slide.test/", "https://slide.test/"),
        ] {
            let xml = xml.replacen("<r>", "<r xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">", 1);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let run = parse_run_with_defaults(doc.root_element(), defaults, &theme, &slide_rels).unwrap();
            assert_eq!(run.hyperlink.as_deref(), Some(click));
            assert_eq!(run.hyperlink_mouse_over.as_deref(), Some(hover));
        }
    }

    #[test]
    fn inherited_slide_jump_is_resolved_from_master_and_layout_parts() {
        let doc = roxmltree::Document::parse(
            r#"<rPr
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>
        </rPr>"#,
        )
        .unwrap();
        let rels = HashMap::from([("rId7".into(), "../slides/slide3.xml".into())]);
        for source_part in [
            "ppt/slideMasters/slideMaster1.xml",
            "ppt/slideLayouts/slideLayout1.xml",
        ] {
            let props = RunProperties::from_xml(doc.root_element(), &HashMap::new())
                .with_relationships(&rels)
                .with_part_targets(source_part);
            assert_eq!(
                inherited_hyperlink_target(
                    &props.hlink_click_target,
                    props.child_attributes.get("hlinkClick")
                ),
                Some("ppt/slides/slide3.xml".into())
            );
        }
        let external = roxmltree::Document::parse(
            r#"<rPr
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <hlinkClick r:id="rId7"/></rPr>"#,
        )
        .unwrap();
        let props = RunProperties::from_xml(external.root_element(), &HashMap::new())
            .with_relationships(&rels)
            .with_part_targets("ppt/slideMasters/slideMaster1.xml");
        assert_eq!(
            inherited_hyperlink_target(
                &props.hlink_click_target,
                props.child_attributes.get("hlinkClick")
            ),
            Some("../slides/slide3.xml".into())
        );
        let nearer = roxmltree::Document::parse(
            r#"<rPr><hlinkClick
          action="ppaction://hlinksldjump"/></rPr>"#,
        )
        .unwrap();
        let merged = RunProperties::from_xml(nearer.root_element(), &HashMap::new()).over(&props);
        assert_eq!(
            inherited_hyperlink_target(
                &merged.hlink_click_target,
                merged.child_attributes.get("hlinkClick")
            ),
            Some("ppt/slides/slide3.xml".into())
        );
    }
}
