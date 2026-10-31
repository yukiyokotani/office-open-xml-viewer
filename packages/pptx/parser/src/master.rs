//! Slide-master / layout inheritance: the per-master extractors (anchors,
//! alignments, ea-line-break, font sizes, per-level sizes/indents/bullets,
//! txStyle bold/italic/colour/spacing, transforms), the layout-placeholder
//! resolver, and the cached `ParsedMaster` (was `MasterBundle`) / `ParsedLayout`
//! bundles. Extracted verbatim from `lib.rs`; the only non-move change is the
//! `MasterBundle` → `ParsedMaster` type rename (fields unchanged).

use crate::fill::{
    parse_background, parse_cust_geom_with_paint, parse_fill, parse_reflection, parse_xfrm,
};
use crate::shape::{
    extract_decorative_shapes, resolve_picture_shape_properties, PictureShapeProperties,
};
use crate::text::{
    complete_level_sizes, empty_level_bullets, extract_level_bullets, extract_level_colors,
    extract_level_faces, extract_level_font_sizes, extract_level_indents,
    extract_level_run_properties_with_rels, has_any_level_bullet, has_any_level_color,
    has_any_level_face, has_any_level_indent, has_any_level_run_properties, has_any_level_size,
    merge_level_bullets, merge_level_colors, merge_level_faces, merge_level_indents,
    merge_level_run_properties, merge_level_sizes, read_level_alignments, read_level_bullets,
    read_level_colors, read_level_defrpr_values, read_level_faces, read_level_font_sizes,
    read_level_indents, read_level_run_properties_with_rels, resolve_latin_face,
    text_property_color, InheritedBodyPr, LevelAlignments, LevelBullets, LevelColors, LevelFaces,
    LevelFontSizes, LevelIndent, LevelIndents, LevelRunProperties, LevelSpacing,
    SharedLevelRunProperties, DEFAULT_TEXT_STYLE_MAR_L, HARD_DEFAULT_FONT_SIZE,
    HARD_DEFAULT_TEXT_COLOR,
};
use crate::theme::{
    bake_clr_map, parse_theme_part, PptxSchemeResolver, PptxTheme, PptxThemeSource,
};
use crate::types::*;
use crate::{
    attr, attr_r, build_smartart_drawings, child, find_rel_target_by_type,
    note_layout_master_parse, parse_preflighted_pptx_xml, parse_rels, read_zip_str, resolve_path,
    PptxZip,
};
use ooxml_common::blip::{
    mime_from_ext, parse_blip_duotone, parse_blip_effects, parse_src_rect, Duotone, SrcRect,
};
use ooxml_common::rels::relationship_part_path;
use std::collections::HashMap;

type MasterTextBodyPropertyMaps = HashMap<String, InheritedBodyPr>;

/// Keyed first by idx (integer), then by type string.
// `Clone` lets `parse_layout` cache one resolved `LayoutPlaceholders` per layout
// and hand each slide a copy to layer its per-slide master txStyles fallbacks
// onto without mutating the cached instance (D4).
#[derive(Default, Clone, serde::Serialize)]
pub(crate) struct LayoutPlaceholders {
    pub(crate) by_idx: HashMap<u32, Transform>,
    /// Effective placeholder type declared by a layout slot. A slide
    /// placeholder may omit @type while retaining @idx; the matching layout
    /// slot supplies its authored type or the CT_Placeholder default (`obj`).
    pub(crate) by_idx_placeholder_type: HashMap<u32, String>,
    pub(crate) by_type: HashMap<String, Transform>,
    /// Fallback transforms from slide master (by ph_type), used when layout has no xfrm
    pub(crate) master_by_type: HashMap<String, Transform>,
    /// Per-list-level Latin typefaces per placeholder idx/type: the layout
    /// placeholder's lstStyle merged per level over the master placeholder and
    /// the master class style. Theme tokens are resolved before storage; a
    /// level no tier names stays `None` and ends at the hard default in
    /// `shape.rs` (issue #1620). The `master` map lets an idx-bearing slide
    /// placeholder with no matching layout slot inherit the master without
    /// borrowing an unrelated layout sibling of the same type.
    pub(crate) by_idx_level_faces: HashMap<u32, LevelFaces>,
    pub(crate) by_type_level_faces: HashMap<String, LevelFaces>,
    pub(crate) by_type_master_level_faces: HashMap<String, LevelFaces>,
    /// Layout placeholder slots that have no `<p:txBody>`, by idx and (for the
    /// first slot of each type, the one the type-keyed maps represent) by type.
    /// A slide placeholder bound to such a slot does not reach the master list
    /// styles; see [`LayoutPlaceholders::is_list_style_cut`].
    pub(crate) by_idx_without_text_body: std::collections::HashSet<u32>,
    pub(crate) by_type_without_text_body: std::collections::HashSet<String>,
    /// The presentation `defaultTextStyle` levels for ordinary (non-placeholder)
    /// text, resolved against this master's theme.
    pub(crate) default_text: DefaultTextLevels,
    /// txStyles-only tier for an idx with no layout slot (`MasterStyleTier`).
    pub(crate) styles: MasterStyleTier,
    /// Per-list-level default font sizes (pt) per placeholder idx — index 0..=8
    /// maps to lvl1pPr..lvl9pPr (ECMA-376 §21.1.2.4). Lets nested bullets shrink
    /// per level (e.g. body 28pt → lvl2 24pt → lvl3 20pt) instead of all using
    /// the level-1 size. None per level where the style chain doesn't specify it.
    pub(crate) by_idx_level_sizes: HashMap<u32, LevelFontSizes>,
    /// Per-list-level default font sizes (pt) per placeholder type.
    pub(crate) by_type_level_sizes: HashMap<String, LevelFontSizes>,
    pub(crate) by_type_master_level_sizes: HashMap<String, LevelFontSizes>,
    pub(crate) by_idx_level_colors: HashMap<u32, LevelColors>,
    pub(crate) by_type_level_colors: HashMap<String, LevelColors>,
    pub(crate) by_type_master_level_colors: HashMap<String, LevelColors>,
    pub(crate) by_idx_level_run_properties: HashMap<u32, SharedLevelRunProperties>,
    pub(crate) by_type_level_run_properties: HashMap<String, SharedLevelRunProperties>,
    /// txStyles-only character properties for an idx with no layout slot
    /// (see `MasterLevelRunProperties`).
    pub(crate) by_type_master_level_run_properties: HashMap<String, SharedLevelRunProperties>,
    /// Per-list-level paragraph indents (`marL`/`marR`/`indent`, EMU) per
    /// placeholder idx — what a paragraph with no own `marL`/`marR`/`indent`
    /// inherits from the authored list-style cascade (ECMA-376 §21.1.2.4.13),
    /// used as the fallback before PowerPoint's hardcoded implicit defaults.
    pub(crate) by_idx_level_indents: HashMap<u32, LevelIndents>,
    /// Per-list-level paragraph indents per placeholder type.
    pub(crate) by_type_level_indents: HashMap<String, LevelIndents>,
    pub(crate) by_type_master_level_indents: HashMap<String, LevelIndents>,
    /// Per-list-level inherited bullet (buChar/buAutoNum/buNone) per placeholder
    /// idx — what a paragraph with no explicit bullet inherits (ECMA-376 §19.7.10).
    pub(crate) by_idx_level_bullets: HashMap<u32, LevelBullets>,
    /// Per-list-level inherited bullet per placeholder type.
    pub(crate) by_type_level_bullets: HashMap<String, LevelBullets>,
    pub(crate) by_type_master_level_bullets: HashMap<String, LevelBullets>,
    /// Default bold per placeholder type, from layout lstStyle defRPr b attribute
    pub(crate) by_type_bold: HashMap<String, bool>,
    /// Default italic per placeholder type, from layout lstStyle defRPr i attribute
    pub(crate) by_type_italic: HashMap<String, bool>,
    /// Default caps ("all"/"small") per placeholder type, from layout/master
    /// lstStyle defRPr cap attribute (ECMA-376 §21.1.2.3.9;
    /// ST_TextCapsType §20.1.10.64)
    pub(crate) by_type_caps: HashMap<String, String>,
    /// Default run reflection per placeholder type, inherited from layout or
    /// master `lvl1pPr/defRPr/effectLst`.
    pub(crate) by_type_reflection: HashMap<String, Reflection>,
    /// Vertical anchor ("t"/"ctr"/"b") per placeholder idx/type, from
    /// layout/master bodyPr. The idx tier prevents one of several same-type
    /// layout slots from leaking its alignment into its siblings.
    pub(crate) by_idx_anchor: HashMap<u32, String>,
    pub(crate) by_type_anchor: HashMap<String, String>,
    pub(crate) by_type_master_anchor: HashMap<String, String>,
    /// Per-placeholder layout `bodyPr` values (insets, wrap, vert, columns,
    /// spcFirstLastPara, autofit child, prstTxWarp), each already merged with
    /// the master value. Every field stays optional so an omitted attribute
    /// continues to the master value and then the schema default instead of
    /// being replaced by a synthetic layout default.
    pub(crate) by_idx_body_pr: HashMap<u32, InheritedBodyPr>,
    pub(crate) by_type_body_pr: HashMap<String, InheritedBodyPr>,
    pub(crate) by_type_master_body_pr: HashMap<String, InheritedBodyPr>,
    /// Default paragraph alignment per placeholder type, from layout/master lstStyle
    pub(crate) by_type_alignment: HashMap<String, String>,
    /// Paragraph alignment per placeholder idx — layout placeholder's own algn,
    /// falling back to the master per-type alignment. Checked before the
    /// type-keyed maps so a body placeholder resolves to its OWN idx's style,
    /// not an unrelated typeless placeholder (ECMA-376 §19.3.1.x idx matching).
    pub(crate) by_idx_alignment: HashMap<u32, String>,
    /// Default East Asian line-break (eaLnBrk) per placeholder type, from the
    /// layout lstStyle > lvl1pPr @eaLnBrk (ECMA-376 §21.1.2.2.7)
    pub(crate) by_type_ea_ln_brk: HashMap<String, bool>,
    /// Default font alignment per placeholder type, from the layout lstStyle >
    /// lvl1pPr @fontAlgn (ECMA-376 §21.1.2.2.7). Same tiers as eaLnBrk; the
    /// master tier lives in `styles` (`MasterStyleTier::placeholder_font_algn`).
    pub(crate) by_type_font_algn: HashMap<String, String>,
    /// fontAlgn per bound layout slot (idx): the slot's own lvl1pPr value,
    /// else the master value for its type, mirroring `by_idx_alignment`.
    pub(crate) by_idx_font_algn: HashMap<u32, String>,
    /// Per-level paragraph spacing (spcBef / spcAft / lnSpc) per placeholder
    /// idx, from the matching layout placeholder's lstStyle. The idx tier
    /// prevents one of several same-type layout slots from leaking paragraph
    /// spacing into its siblings (ECMA-376 §19.3.1.36 placeholder matching).
    pub(crate) by_idx_spacing: HashMap<u32, LevelSpacing>,
    /// Per-level paragraph spacing per placeholder type, from layout lstStyle
    pub(crate) by_type_spacing: HashMap<String, LevelSpacing>,
    /// Per-level paragraph spacing from the master placeholders and txStyles
    /// (fallback per level when the layout has none)
    pub(crate) by_type_master_spacing: HashMap<String, LevelSpacing>,
    /// Stroke per placeholder type from layout spPr > ln
    pub(crate) by_type_stroke: HashMap<String, Stroke>,
    /// Stroke per placeholder idx from layout spPr > ln
    pub(crate) by_idx_stroke: HashMap<u32, Stroke>,
    /// Complete picture-like shape properties inherited from the matching
    /// layout placeholder. This accompanies an inherited `blipFill` so every
    /// PictureElement construction path retains the same effects and 3-D
    /// components as an ordinary `p:pic`.
    pub(crate) by_type_picture_properties: HashMap<String, PictureShapeProperties>,
    pub(crate) by_idx_picture_properties: HashMap<u32, PictureShapeProperties>,
    /// Paragraph alignment per placeholder type from master lstStyle > lvl1pPr algn (fallback)
    pub(crate) by_type_master_alignment: HashMap<String, String>,
    /// East Asian line-break per placeholder type from master lstStyle > lvl1pPr
    /// @eaLnBrk (fallback when the layout has none) — ECMA-376 §21.1.2.2.7
    pub(crate) by_type_master_ea_ln_brk: HashMap<String, bool>,
    /// Inherited blipFill (data URL + src rect) per placeholder idx from layout spPr
    pub(crate) by_idx_blip_fill: HashMap<u32, InheritedBlipFill>,
    /// Inherited blipFill per placeholder type from layout spPr
    pub(crate) by_type_blip_fill: HashMap<String, InheritedBlipFill>,
    /// Default text color per placeholder idx, from layout lstStyle defRPr solidFill
    pub(crate) by_idx_color: HashMap<u32, String>,
    /// Default text color per placeholder type, from layout lstStyle defRPr solidFill
    pub(crate) by_type_color: HashMap<String, String>,
    /// Default text color from master (txStyles + spTree lstStyle) — fallback when layout has none
    pub(crate) by_type_master_color: HashMap<String, String>,
    /// `<p:spPr><a:solidFill | a:noFill | a:gradFill | a:pattFill>` per placeholder idx.
    /// Used to inherit a layout-level shape fill (e.g. a tinted body placeholder)
    /// onto slide-level shapes whose `<p:spPr>` is empty.
    pub(crate) by_idx_fill: HashMap<u32, Fill>,
    /// Same as `by_idx_fill` but keyed by placeholder type (fallback when idx
    /// doesn't match a layout shape).
    pub(crate) by_type_fill: HashMap<String, Fill>,
    /// Shape geometry from the matching layout placeholder. Presentation slides
    /// inherit layout information unless they provide a local override
    /// (ECMA-376 Part 1, Annex L.3.2.3). This includes the preset/custom
    /// geometry and preset adjustment values, not just the transform and paint.
    pub(crate) by_idx_geometry: HashMap<u32, InheritedShapeGeometry>,
    /// Type-keyed geometry fallback for placeholders that do not declare `idx`.
    pub(crate) by_type_geometry: HashMap<String, InheritedShapeGeometry>,
}

#[derive(Debug, Clone, serde::Serialize)]
pub(crate) struct InheritedShapeGeometry {
    pub(crate) geometry: String,
    pub(crate) cust_geom: Option<Vec<Vec<PathCmd>>>,
    pub(crate) cust_geom_paint: Option<Vec<PathPaint>>,
    pub(crate) adjustments: [Option<f64>; 8],
}

impl InheritedShapeGeometry {
    /// Parse the geometry-bearing portion of `<p:spPr>`. `None` means the shape
    /// did not locally specify geometry and therefore remains eligible for
    /// placeholder inheritance.
    pub(crate) fn from_sp_pr(
        sp_pr: roxmltree::Node<'_, '_>,
        shape_w: f64,
        shape_h: f64,
    ) -> Option<Self> {
        let cust_geom_node = child(sp_pr, "custGeom");
        let prst_geom_node = child(sp_pr, "prstGeom");
        if let Some(cust_geom_node) = cust_geom_node {
            let (paths, paint) = parse_cust_geom_with_paint(cust_geom_node, shape_w, shape_h);
            return Some(Self {
                geometry: "custGeom".to_owned(),
                cust_geom: Some(paths),
                cust_geom_paint: paint,
                adjustments: [None; 8],
            });
        }

        let prst_geom_node = prst_geom_node?;
        let geometry = attr(&prst_geom_node, "prst")?;
        let gd_nodes: Vec<_> = child(prst_geom_node, "avLst")
            .map(|av| {
                av.children()
                    .filter(|n| n.is_element() && n.tag_name().name() == "gd")
                    .collect()
            })
            .unwrap_or_default();
        Some(Self {
            geometry,
            cust_geom: None,
            cust_geom_paint: None,
            adjustments: std::array::from_fn(|index| preset_adjustment(&gd_nodes, index)),
        })
    }
}

/// Adjust handle `index` (0 = `adj`/`adj1`) from a `<a:prstGeom><a:avLst>`'s
/// `<a:gd>` guides, as PowerPoint shapes and connectors read it.
///
/// This intentionally differs from the DOCX/XLSX
/// `ooxml_common::preset_geometry::parse_preset_adj`: when no guide carries
/// the handle's name, the guide at the same position is used even in a named
/// list, and the `val` operand is not trimmed. Changing it would move shapes
/// in decks that author only some of a multi-handle preset's guides.
pub(crate) fn preset_adjustment(gd_nodes: &[roxmltree::Node<'_, '_>], index: usize) -> Option<f64> {
    let expected_name = if index == 0 {
        None
    } else {
        Some(format!("adj{}", index + 1))
    };
    gd_nodes
        .iter()
        .find(|n| {
            let name = attr(n, "name");
            if index == 0 {
                matches!(name.as_deref(), Some("adj") | Some("adj1"))
            } else {
                name == expected_name
            }
        })
        .or_else(|| gd_nodes.get(index))
        .and_then(|gd| attr(gd, "fmla"))
        .and_then(|fmla| fmla.strip_prefix("val ").map(str::to_owned))
        .and_then(|value| value.parse::<f64>().ok())
}

#[derive(Debug, Clone, serde::Serialize)]
pub(crate) struct InheritedBlipFill {
    /// Embedded zip path of the inherited picture-placeholder blip.
    pub(crate) image_path: String,
    /// MIME of the blip at `image_path`.
    pub(crate) mime_type: String,
    pub(crate) src_rect: Option<SrcRect>,
    pub(crate) alpha: Option<f64>,
    /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour on the layout placeholder's
    /// blipFill, resolved through the theme. Inherited onto the slide picture
    /// placeholder that omits its own blipFill (see `shape.rs`).
    pub(crate) duotone: Option<Duotone>,
    /// CT_Blip pixel effects on the same blipFill, inherited with it.
    pub(crate) blip_effects: Vec<ooxml_common::blip::BlipEffect>,
}

impl LayoutPlaceholders {
    pub(crate) fn lookup(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<&Transform> {
        ph_idx
            .and_then(|i| self.by_idx.get(&i))
            .or_else(|| self.by_type.get(ph_type))
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type.get("")
                } else {
                    None
                }
            })
            .or_else(|| self.master_by_type.get(ph_type))
            // ECMA-376 §19.7.9 defines placeholder size relative to the body
            // placeholder on the master. An object/content layout slot with no
            // own transform therefore uses the master body box; its semantic
            // placeholder type remains the CT_Placeholder default (`obj`).
            .or_else(|| {
                if ph_type == "obj" {
                    self.master_by_type.get("body")
                } else {
                    None
                }
            })
    }

    /// A slide placeholder whose matched layout slot has no `<p:txBody>` does
    /// not inherit the master placeholder or master text styles. Observed with
    /// PowerPoint for Mac PDF export (issue #1620): title, body, obj,
    /// subTitle, dt, ftr and sldNum slots without a txBody rendered their text
    /// in Arial 18 pt with no bullet, on masters with and without txStyles, while
    /// the same placeholders bound to a slot WITH a txBody (even an empty one)
    /// and placeholders with no matching slot at all used the master styles.
    /// The unit that stops is the list style, so every list-style property is
    /// cut, not only the observed face, size and bullet.
    pub(crate) fn is_list_style_cut(&self, ph_type: &str, ph_idx: Option<u32>) -> bool {
        if let Some(i) = ph_idx {
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self.by_idx_without_text_body.contains(&i);
            }
            return false;
        }
        self.by_type_without_text_body.contains(ph_type)
    }

    /// Per-list-level inherited Latin faces. Same idx-strict resolution as
    /// `lookup_level_font_sizes`.
    pub(crate) fn lookup_level_faces(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelFaces {
        if let Some(i) = ph_idx {
            // Bound slot: layout -> master placeholder -> class style,
            // already merged. Unmatched idx: the style tier below.
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self.by_idx_level_faces.get(&i).cloned().unwrap_or_default();
            }
            return self
                .by_idx_level_faces
                .get(&i)
                .cloned()
                .or_else(|| self.by_type_master_level_faces.get(ph_type).cloned())
                .or_else(|| {
                    if ph_type == "obj" {
                        self.by_type_master_level_faces.get("").cloned()
                    } else {
                        None
                    }
                })
                .unwrap_or_default();
        }
        self.by_type_level_faces
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_faces.get("").cloned()
                } else {
                    None
                }
            })
            .unwrap_or_default()
    }

    /// Per-list-level inherited default font sizes (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_level_faces`. All-None when the placeholder has no
    /// per-level styling.
    pub(crate) fn lookup_level_font_sizes(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> LevelFontSizes {
        if let Some(i) = ph_idx {
            // Bound slot: layout -> master placeholder -> class style,
            // already merged. Unmatched idx: the style tier below.
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self
                    .by_idx_level_sizes
                    .get(&i)
                    .copied()
                    .unwrap_or([None; 9]);
            }
            return self
                .by_idx_level_sizes
                .get(&i)
                .copied()
                .or_else(|| self.by_type_master_level_sizes.get(ph_type).copied())
                .or_else(|| {
                    if ph_type == "obj" {
                        self.by_type_master_level_sizes.get("").copied()
                    } else {
                        None
                    }
                })
                .unwrap_or([None; 9]);
        }
        self.by_type_level_sizes
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_sizes.get("").copied()
                } else {
                    None
                }
            })
            .unwrap_or([None; 9])
    }

    pub(crate) fn lookup_level_colors(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelColors {
        if let Some(i) = ph_idx {
            // Bound slot: layout -> master placeholder -> class style,
            // already merged. Unmatched idx: the style tier below.
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self
                    .by_idx_level_colors
                    .get(&i)
                    .cloned()
                    .unwrap_or_else(|| std::array::from_fn(|_| None));
            }
            return self
                .by_idx_level_colors
                .get(&i)
                .cloned()
                .or_else(|| self.by_type_master_level_colors.get(ph_type).cloned())
                .or_else(|| {
                    if ph_type == "obj" {
                        self.by_type_master_level_colors.get("").cloned()
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| std::array::from_fn(|_| None));
        }
        self.by_type_level_colors
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_colors.get("").cloned()
                } else {
                    None
                }
            })
            .unwrap_or_else(|| std::array::from_fn(|_| None))
    }

    pub(crate) fn lookup_level_run_properties(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> LevelRunProperties {
        let empty = || std::array::from_fn(|_| Default::default());
        if let Some(i) = ph_idx {
            // A bound layout slot carries layout ∪ master placeholder ∪ txStyles;
            // an idx with no slot reaches txStyles only (MasterLevelRunProperties).
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self
                    .by_idx_level_run_properties
                    .get(&i)
                    .map(|levels| LevelRunProperties::clone(levels))
                    .unwrap_or_else(empty);
            }
            return self
                .by_type_master_level_run_properties
                .get(ph_type)
                .map(|levels| LevelRunProperties::clone(levels))
                .or_else(|| {
                    (ph_type == "obj")
                        .then(|| {
                            self.by_type_master_level_run_properties
                                .get("")
                                .map(|levels| LevelRunProperties::clone(levels))
                        })
                        .flatten()
                })
                .unwrap_or_else(empty);
        }
        self.by_type_level_run_properties
            .get(ph_type)
            .map(|levels| LevelRunProperties::clone(levels))
            .or_else(|| {
                (ph_type == "body")
                    .then(|| {
                        self.by_type_level_run_properties
                            .get("")
                            .map(|levels| LevelRunProperties::clone(levels))
                    })
                    .flatten()
            })
            .unwrap_or_else(empty)
    }

    /// Per-list-level inherited paragraph indents (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_level_font_sizes`. All-default (every axis None) when
    /// the placeholder has no authored per-level indent.
    pub(crate) fn lookup_level_indents(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelIndents {
        if let Some(i) = ph_idx {
            // Bound slot: layout -> master placeholder -> class style,
            // already merged. Unmatched idx: the style tier below.
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self
                    .by_idx_level_indents
                    .get(&i)
                    .copied()
                    .unwrap_or_default();
            }
            return self
                .by_idx_level_indents
                .get(&i)
                .copied()
                .or_else(|| self.by_type_master_level_indents.get(ph_type).copied())
                .or_else(|| {
                    if ph_type == "obj" {
                        self.by_type_master_level_indents.get("").copied()
                    } else {
                        None
                    }
                })
                .unwrap_or_default();
        }
        self.by_type_level_indents
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_indents.get("").copied()
                } else {
                    None
                }
            })
            .unwrap_or_default()
    }

    /// Per-list-level inherited bullets (lvl1..lvl9). Same idx-strict resolution as
    /// `lookup_level_font_sizes`. All-None when the placeholder inherits no bullet.
    pub(crate) fn lookup_level_bullets(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelBullets {
        if let Some(i) = ph_idx {
            // Bound slot: layout -> master placeholder -> class style,
            // already merged. Unmatched idx: the style tier below.
            if self.by_idx_placeholder_type.contains_key(&i) {
                return self
                    .by_idx_level_bullets
                    .get(&i)
                    .cloned()
                    .unwrap_or_else(empty_level_bullets);
            }
            return self
                .by_idx_level_bullets
                .get(&i)
                .cloned()
                .or_else(|| self.by_type_master_level_bullets.get(ph_type).cloned())
                .or_else(|| {
                    if ph_type == "obj" {
                        self.by_type_master_level_bullets.get("").cloned()
                    } else {
                        None
                    }
                })
                .unwrap_or_else(empty_level_bullets);
        }
        self.by_type_level_bullets
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_level_bullets.get("").cloned()
                } else {
                    None
                }
            })
            .unwrap_or_else(empty_level_bullets)
    }

    /// Look up inherited bold for this placeholder type.
    pub(crate) fn lookup_bold(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<bool> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.bold, ph_type).cloned();
            }
        }
        self.by_type_bold.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_bold.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited italic for this placeholder type.
    pub(crate) fn lookup_italic(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<bool> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.italic, ph_type).cloned();
            }
        }
        self.by_type_italic.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_italic.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited caps ("all"/"small") for this placeholder type.
    pub(crate) fn lookup_caps(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.caps, ph_type).cloned();
            }
        }
        self.by_type_caps.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_caps.get("").cloned()
            } else {
                None
            }
        })
    }

    pub(crate) fn lookup_reflection(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<Reflection> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.reflection, ph_type).cloned();
            }
        }
        self.by_type_reflection.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_reflection.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up inherited vertical anchor for this placeholder. An anchor on the
    /// exact idx-matched layout slot wins. PowerPoint otherwise retains the
    /// layout's type-level placeholder fallback before consulting the master;
    /// this preserves layouts whose first same-type slot carries the shared
    /// anchor while still preventing it from overriding an explicitly authored
    /// anchor on a later idx.
    pub(crate) fn lookup_anchor(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        let master = || {
            self.by_type_master_anchor
                .get(ph_type)
                .cloned()
                .or_else(|| {
                    if ph_type == "body" || ph_type == "obj" {
                        self.by_type_master_anchor.get("").cloned()
                    } else {
                        None
                    }
                })
        };
        if let Some(i) = ph_idx {
            return self
                .by_idx_anchor
                .get(&i)
                .cloned()
                .or_else(|| self.by_type_anchor.get(ph_type).cloned())
                .or_else(|| {
                    if ph_type == "body" || ph_type == "obj" {
                        self.by_type_anchor.get("").cloned()
                    } else {
                        None
                    }
                })
                .or_else(master);
        }
        self.by_type_anchor.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" || ph_type == "obj" {
                self.by_type_anchor.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up the layout/master `bodyPr` values for this placeholder. An
    /// explicit idx is strict so a body placeholder cannot borrow another body
    /// slot's properties.
    pub(crate) fn lookup_body_pr(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<InheritedBodyPr> {
        if let Some(i) = ph_idx {
            return self
                .by_idx_body_pr
                .get(&i)
                .cloned()
                .or_else(|| self.by_type_master_body_pr.get(ph_type).cloned());
        }
        self.by_type_body_pr
            .get(ph_type)
            .cloned()
            .or_else(|| self.by_type_master_body_pr.get(ph_type).cloned())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_body_pr
                        .get("")
                        .cloned()
                        .or_else(|| self.by_type_master_body_pr.get("").cloned())
                } else {
                    None
                }
            })
    }

    /// Look up inherited paragraph alignment for this placeholder.
    ///
    /// A placeholder identified by `idx` resolves through its own slot
    /// (`by_idx_alignment`), which `parse_layout_placeholders` pre-seeds with the
    /// master per-type default. Unlike `lookup_fill`, falling through to the
    /// type map on an idx miss is intentional and safe (the seed already encodes
    /// the master tier) — but the `""` (typeless) fallback is gated to
    /// `ph_idx.is_none()` so an idx/typed placeholder never borrows an unrelated
    /// typeless sibling's alignment (ECMA-376 §19.3.1.36 idx matching).
    pub(crate) fn lookup_alignment(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.alignment, ph_type).cloned();
            }
        }
        if let Some(i) = ph_idx {
            if let Some(a) = self.by_idx_alignment.get(&i) {
                return Some(a.clone());
            }
        }
        // The `""` fallback represents a typeless (idx-less, body-category)
        // placeholder; only a placeholder that is itself typeless may use it.
        let allow_empty = ph_idx.is_none() && ph_type == "body";
        self.by_type_alignment
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if allow_empty {
                    self.by_type_alignment.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_alignment.get(ph_type).cloned())
            .or_else(|| {
                if allow_empty {
                    self.by_type_master_alignment.get("").cloned()
                } else {
                    None
                }
            })
    }

    // ECMA-376 §21.1.2.2.7 eaLnBrk inheritance, mirroring lookup_alignment:
    // layout per-type → layout generic ("") for body → master per-type →
    // master generic. None means no ancestor specified it (parse_paragraph then
    // applies the spec default of true).
    pub(crate) fn lookup_ea_ln_brk(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<bool> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.ea_ln_brk, ph_type).cloned();
            }
        }
        self.by_type_ea_ln_brk
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_ea_ln_brk.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_ea_ln_brk.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_ea_ln_brk.get("").copied()
                } else {
                    None
                }
            })
    }

    // ECMA-376 §21.1.2.2.7 fontAlgn inheritance, mirroring lookup_ea_ln_brk.
    // A placeholder bound to a layout slot by idx reads exactly that slot,
    // then the master placeholder of the slot's class, then txStyles; it never
    // reads a same-type sibling slot. An idx without a layout slot reads the
    // class list style only. Only an idx-less placeholder uses the type maps.
    pub(crate) fn lookup_font_algn(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            return match self.by_idx_placeholder_type.get(&i) {
                None => MasterStyleTier::get(&self.styles.font_algn, ph_type).cloned(),
                Some(slot_type) => self.by_idx_font_algn.get(&i).cloned().or_else(|| {
                    MasterStyleTier::get(&self.styles.placeholder_font_algn, slot_type).cloned()
                }),
            };
        }
        let master = &self.styles.placeholder_font_algn;
        self.by_type_font_algn
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_font_algn.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| master.get(ph_type).cloned())
            .or_else(|| {
                if ph_type == "body" {
                    master.get("").cloned()
                } else {
                    None
                }
            })
    }

    /// Per-level paragraph spacing for this placeholder: the layout slot (by
    /// idx when the placeholder has one, else by type), then the master, per
    /// level and property. An idx with no layout slot takes the txStyles-only
    /// tier. The layout placeholder itself inherits the master text style for
    /// its type (ECMA-376 §19.3.1.36 / §19.3.1.51), so an idx-matched slot
    /// without spacing still yields the master values.
    pub(crate) fn lookup_spacing(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelSpacing {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.spacing, ph_type)
                    .cloned()
                    .unwrap_or_default();
            }
        }
        let by_type = |map: &HashMap<String, LevelSpacing>| -> LevelSpacing {
            let own = map.get(ph_type).cloned().unwrap_or_default();
            if ph_type == "body" {
                own.or(&map.get("").cloned().unwrap_or_default())
            } else {
                own
            }
        };
        let layout = match ph_idx {
            Some(idx) => self.by_idx_spacing.get(&idx).cloned().unwrap_or_default(),
            None => by_type(&self.by_type_spacing),
        };
        layout.or(&by_type(&self.by_type_master_spacing))
    }

    /// Look up inherited blipFill from the layout placeholder spPr. Used when a slide
    /// references a picture placeholder (e.g. ph type="pic") without its own blipFill —
    /// the image defined on the layout's matching placeholder should render through.
    /// Idx-strict per ECMA-376 §19.3.1.36 (see `lookup_fill`'s rationale).
    pub(crate) fn lookup_blip_fill(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<InheritedBlipFill> {
        if let Some(i) = ph_idx {
            return self.by_idx_blip_fill.get(&i).cloned();
        }
        self.by_type_blip_fill.get(ph_type).cloned()
    }

    /// Look up inherited stroke from the layout placeholder spPr > ln.
    /// Idx-strict per ECMA-376 §19.3.1.36 (see `lookup_fill`'s rationale).
    pub(crate) fn lookup_stroke(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<Stroke> {
        if let Some(i) = ph_idx {
            return self.by_idx_stroke.get(&i).cloned();
        }
        self.by_type_stroke.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_stroke.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up all picture-affecting shape properties from the same placeholder
    /// slot as an inherited blipFill. Explicit idx matching remains strict per
    /// §19.3.1.36, exactly like fill, geometry, stroke and blipFill.
    pub(crate) fn lookup_picture_properties(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<PictureShapeProperties> {
        if let Some(i) = ph_idx {
            return self.by_idx_picture_properties.get(&i).cloned().or_else(|| {
                self.by_idx_stroke
                    .get(&i)
                    .cloned()
                    .map(|stroke| PictureShapeProperties {
                        stroke: Some(stroke),
                        ..Default::default()
                    })
            });
        }
        self.by_type_picture_properties
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_picture_properties.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| {
                self.lookup_stroke(ph_type, None)
                    .map(|stroke| PictureShapeProperties {
                        stroke: Some(stroke),
                        ..Default::default()
                    })
            })
    }

    /// Look up inherited default text color for this placeholder (layout then master fallback).
    ///
    /// The *layout* tier is idx-strict per ECMA-376 §19.3.1.36: when the slide-level
    /// placeholder carries an explicit `idx`, a layout colour is inherited only from the
    /// layout shape with the SAME idx — never a sibling body placeholder at a different
    /// idx (which would leak an unrelated region's colour).
    ///
    /// The *master* class-style tier (titleStyle/bodyStyle/defaultTextStyle), however, is a
    /// document-wide default keyed by placeholder *type* (§21.1.2.4 / §19.3.1) and is
    /// inherited regardless of idx. So when the idx-matched layout shape defines no
    /// colour, resolution must still fall through to `by_type_master_color`. Without
    /// this, a body placeholder whose layout shape sets size-but-not-colour resolves to
    /// no colour at all and the renderer defaults to black — instead of the master
    /// bodyStyle colour (e.g. `schemeClr bg1` = white on a dark theme). (sample-9 slide 2+)
    pub(crate) fn lookup_color(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            if !self.by_idx_placeholder_type.contains_key(&i) {
                return MasterStyleTier::get(&self.styles.color, ph_type).cloned();
            }
        }
        if let Some(i) = ph_idx {
            if let Some(c) = self.by_idx_color.get(&i) {
                return Some(c.clone());
            }
            // Layout idx had no colour → fall through to the master type-keyed default.
            return self.by_type_master_color.get(ph_type).cloned().or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_color.get("").cloned()
                } else {
                    None
                }
            });
        }
        self.by_type_color
            .get(ph_type)
            .cloned()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_color.get("").cloned()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_color.get(ph_type).cloned())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_color.get("").cloned()
                } else {
                    None
                }
            })
    }

    /// Look up the inherited shape fill from the layout placeholder's `<p:spPr>`.
    /// Used when the slide-level shape leaves `<p:spPr>` empty (or with no fill
    /// elements) and is bound to a placeholder.
    ///
    /// ECMA-376 §19.3.1.36 (placeholder inheritance) is asymmetric: when the
    /// slide-level shape declares `<p:ph idx="N">` it is bound to *that*
    /// specific layout slot — the only valid inheritance source is the layout
    /// shape with idx=N. Falling back to `by_type_fill` here would let a
    /// sibling body placeholder (a different idx, different region of the
    /// layout) bleed its fill onto a placeholder that the spec says should
    /// have no fill. This is exactly what regressed sample-2 slide-4: layout10
    /// has `body[idx=12]` (header, no fill) and `body[idx=13]` (bullet box,
    /// gray fill) — the type fallback was leaking the bullet box's gray onto
    /// the header.
    ///
    /// The type-only fallback only applies when the slide-level shape itself
    /// has no idx, in which case "first body placeholder we found" is the
    /// best we can do.
    pub(crate) fn lookup_fill(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<Fill> {
        if let Some(i) = ph_idx {
            return self.by_idx_fill.get(&i).cloned();
        }
        self.by_type_fill.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_fill.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Look up geometry from the matching layout placeholder. Like fill,
    /// stroke, and blipFill, an explicit `idx` is strict: a slide placeholder
    /// must never borrow geometry from a different body slot merely because
    /// their placeholder types happen to match.
    pub(crate) fn lookup_geometry(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<InheritedShapeGeometry> {
        if let Some(i) = ph_idx {
            return self.by_idx_geometry.get(&i).cloned();
        }
        self.by_type_geometry.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_geometry.get("").cloned()
            } else {
                None
            }
        })
    }
}

/// Parse bodyPr anchor ("t"/"ctr"/"b") from master placeholder shapes.
///
/// Takes the already-parsed master root element (`<p:sldMaster>`) so
/// `build_master_bundle` can parse the master XML once and share the
/// `Document` across every `parse_master_*` extractor (ECMA-376 §19.3.1.42).
pub(crate) fn parse_master_anchors(root: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(anchor) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "bodyPr"))
                    .and_then(|bp| attr(&bp, "anchor"))
                {
                    map.entry(ph_type).or_insert(anchor.to_string());
                }
            }
        }
    }
    map
}

/// Parse the placeholder-scoped `bodyPr` values that must survive the master →
/// layout → slide cascade even when the layout/slide authors an empty bodyPr.
pub(crate) fn parse_master_text_body_properties(
    root: roxmltree::Node<'_, '_>,
) -> MasterTextBodyPropertyMaps {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            else {
                continue;
            };
            let ph_type = attr(&ph, "type").unwrap_or_default();
            let Some(body_pr) = child(sp, "txBody").and_then(|tb| child(tb, "bodyPr")) else {
                continue;
            };
            let value = InheritedBodyPr::from_body_pr(body_pr);
            if !value.is_empty() {
                // Same-type master placeholders merge field by field: the
                // first placeholder in document order that sets a field wins.
                let merged = match map.remove(&ph_type) {
                    Some(first) => InheritedBodyPr::or(first, &value),
                    None => value,
                };
                map.insert(ph_type, merged);
            }
        }
    }
    map
}

/// Placeholder classes and the list style whose Latin FACE and SIZE each one
/// inherits after the layout and master placeholders (ECMA-376 §19.3.1.52
/// txStyles, §19.2.1.8 defaultTextStyle). Every other list-style property
/// keeps the txStyles mapping of `tx_style_nodes`.
///
/// * title / ctrTitle → master `titleStyle` (§19.3.1.49);
/// * body / subTitle / obj / typeless → master `bodyStyle` (§19.3.1.5);
/// * dt / ftr / sldNum → the presentation `defaultTextStyle`.
///
/// The last row is observed PowerPoint behaviour, not §19.3.1.35: with a
/// master whose `otherStyle` named Trebuchet MS 16 pt (lvl1) and Gabriola
/// 15 pt (lvl2), PowerPoint for Mac rendered date, footer and slide-number
/// placeholders in the `defaultTextStyle` face and size of the paragraph's
/// level (Century Gothic 20 pt; lvl2, which named no face, in Arial 20 pt),
/// on masters with and without txStyles (issue #1620). Non-placeholder text
/// on the slide, layout and master used `defaultTextStyle` as well, so no
/// observed face or size came from `otherStyle`. Only face and size were
/// observable in those controls, so the switch is limited to them.
pub(crate) const TITLE_CLASS: &[&str] = &["title", "ctrTitle"];
pub(crate) const BODY_CLASS: &[&str] = &["body", "subTitle", "obj", ""];
pub(crate) const OTHER_CLASS: &[&str] = &["dt", "ftr", "sldNum"];

/// The class list styles that supply the Latin face and size: `titleStyle` /
/// `bodyStyle` from the master's effective txStyles (`effective_tx_styles`)
/// and the presentation `defaultTextStyle` for the other class.
pub(crate) fn class_style_nodes<'a, 'i>(
    root: roxmltree::Node<'a, 'i>,
    default_text_style: Option<roxmltree::Node<'a, 'i>>,
) -> Vec<(roxmltree::Node<'a, 'i>, &'static [&'static str])> {
    let mut out = Vec::new();
    if let Some(tx_styles) = effective_tx_styles(root) {
        if let Some(n) = child(tx_styles, "titleStyle") {
            out.push((n, TITLE_CLASS));
        }
        if let Some(n) = child(tx_styles, "bodyStyle") {
            out.push((n, BODY_CLASS));
        }
    }
    if let Some(n) = default_text_style {
        out.push((n, OTHER_CLASS));
    }
    out
}

/// The master txStyles class mapping for every list-style property other than
/// the Latin face and size: title/ctrTitle → titleStyle, body/subTitle/obj →
/// bodyStyle, dt/ftr/sldNum → otherStyle (ECMA-376 §19.3.1.52). The #1620
/// controls only evidence the face and size of the other class (which follow
/// `class_style_nodes`), so the other properties keep this mapping.
pub(crate) fn tx_style_nodes<'a, 'i>(
    root: roxmltree::Node<'a, 'i>,
) -> Vec<(roxmltree::Node<'a, 'i>, &'static [&'static str])> {
    let Some(tx_styles) = effective_tx_styles(root) else {
        return Vec::new();
    };
    [
        ("titleStyle", TITLE_CLASS),
        ("bodyStyle", BODY_CLASS),
        ("otherStyle", OTHER_CLASS),
    ]
    .into_iter()
    .filter_map(|(name, types)| child(tx_styles, name).map(|n| (n, types)))
    .collect()
}

/// The title and body styles PowerPoint applies to a master without
/// `p:txStyles`, which CT_SlideMaster makes optional (ECMA-376 §19.3.1.42).
///
/// Observed (issues #1435, #1620, #1630; PowerPoint's reference PDF export):
/// on masters without txStyles - with and without master placeholders, under
/// two themes - every measured title and body paragraph at levels 1-5 was laid
/// out exactly like the same text on a master whose txStyles spelled out:
/// * title level 1: theme major Latin, 44 pt, line spacing 90 %, no space
///   before, no bullet. Deeper title levels have no entry: they rendered in
///   Arial 18 pt with single spacing and no indent, the hard defaults;
/// * body levels 1-5: theme minor Latin at 28 / 24 / 20 / 18 / 18 pt, line
///   spacing 90 %, space before 10 pt at level 1 and 5 pt below, an Arial
///   U+2022 bullet, marL 0.25" plus 0.5" per level with a 0.25" hanging
///   indent.
///
/// * colour: title, subTitle and body text with no fill of its own rendered
///   the theme tx1 through the slide's colour map (a dk1 of 10263F against
///   the 000000 hard default), the default template's `schemeClr tx1` fill.
///
/// ctrTitle takes the title style and obj / subTitle / typeless the body style
/// (`class_style_nodes`). These are the values of PowerPoint's default
/// template. A supplementary control measured levels 6-9 the same way: body
/// levels 6-9 continue at marL + 0.5" per level, 18 pt, 90 %, 5 pt before
/// with the same bullet, and title levels 6-9 stay at the hard defaults; both
/// laid out identically to a master with the template txStyles spelled out.
/// No otherStyle is synthesized: no control observed one.
const BUILT_IN_TX_STYLES: &str = concat!(
    r#"<p:txStyles xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
    r#"<p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc>"#,
    r#"<a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>"#,
    r#"<a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>"#,
    r#"<p:bodyStyle>"#,
    r#"<a:lvl1pPr marL="228600" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>"#,
    r#"<a:lvl2pPr marL="685800" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl2pPr>"#,
    r#"<a:lvl3pPr marL="1143000" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl3pPr>"#,
    r#"<a:lvl4pPr marL="1600200" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl4pPr>"#,
    r#"<a:lvl5pPr marL="2057400" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl5pPr>"#,
    r#"<a:lvl6pPr marL="2514600" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl6pPr>"#,
    r#"<a:lvl7pPr marL="2971800" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl7pPr>"#,
    r#"<a:lvl8pPr marL="3429000" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl8pPr>"#,
    r#"<a:lvl9pPr marL="3886200" indent="-228600"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl9pPr>"#,
    r#"</p:bodyStyle></p:txStyles>"#,
);

static BUILT_IN_TX_STYLES_DOC: std::sync::LazyLock<roxmltree::Document<'static>> =
    std::sync::LazyLock::new(|| {
        roxmltree::Document::parse(BUILT_IN_TX_STYLES).expect("built-in txStyles is well-formed")
    });

/// The master's `p:txStyles`, or `BUILT_IN_TX_STYLES` when it has none.
pub(crate) fn effective_tx_styles<'a, 'i>(
    root: roxmltree::Node<'a, 'i>,
) -> Option<roxmltree::Node<'a, 'i>> {
    child(root, "txStyles").or_else(|| Some(BUILT_IN_TX_STYLES_DOC.root_element()))
}

/// A master placeholder's `txBody/lstStyle`.
fn placeholder_list_style<'a, 'i>(sp: roxmltree::Node<'a, 'i>) -> Option<roxmltree::Node<'a, 'i>> {
    child(sp, "txBody").and_then(|tb| child(tb, "lstStyle"))
}

/// A master placeholder's `txBody/lstStyle/lvl1pPr`.
fn master_placeholder_lvl1<'a, 'i>(sp: roxmltree::Node<'a, 'i>) -> Option<roxmltree::Node<'a, 'i>> {
    child(sp, "txBody")
        .and_then(|tb| child(tb, "lstStyle"))
        .and_then(|ls| child(ls, "lvl1pPr"))
}

/// Master placeholder shapes as (effective type, shape) in document order.
fn master_placeholder_shapes<'a, 'i>(
    root: roxmltree::Node<'a, 'i>,
    with_placeholders: bool,
) -> Vec<(String, roxmltree::Node<'a, 'i>)> {
    if !with_placeholders {
        return Vec::new();
    }
    let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) else {
        return Vec::new();
    };
    sp_tree
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        .filter_map(|sp| {
            let ph = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")?;
            Some((attr(&ph, "type").unwrap_or_default(), sp))
        })
        .collect()
}

/// A master normally carries one title and one body placeholder. A layout or
/// slide placeholder of another type in the same class inherits that master
/// placeholder, as placeholder boxes already do (`LayoutPlaceholders::lookup`,
/// obj → master body per §19.7.9). Observed for text (issue #1620): an obj
/// slot took the master body placeholder's face and its buNone, not the
/// bodyStyle face and bullet. Applied to the face, size and paragraph-level
/// maps (bullets, indents, alignment, eaLnBrk); the other character
/// properties keep the per-type master placeholder lookup of #1625.
const MASTER_PLACEHOLDER_CLASSES: [(&str, &[&str]); 2] =
    [("title", &["ctrTitle"]), ("body", &["subTitle", "obj", ""])];

fn inherit_master_placeholder_classes<T: Clone>(map: &mut HashMap<String, T>) {
    for (source, targets) in MASTER_PLACEHOLDER_CLASSES {
        if let Some(value) = map.get(source).cloned() {
            for target in targets {
                map.entry((*target).to_owned())
                    .or_insert_with(|| value.clone());
            }
        }
    }
}

/// Parse paragraph alignment from master placeholder shapes' lstStyle > lvl1pPr algn,
/// then the class list styles.
pub(crate) fn parse_master_alignments(root: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
    parse_master_alignments_tier(root, true)
}

pub(crate) fn parse_master_alignments_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(algn) = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"))
            .and_then(|lp| attr(&lp, "algn"))
        {
            map.entry(ph_type).or_insert(algn);
        }
    }
    inherit_master_placeholder_classes(&mut map);
    for (style, types) in tx_style_nodes(root) {
        if let Some(algn) = child(style, "lvl1pPr").and_then(|lp| attr(&lp, "algn")) {
            for t in types {
                map.entry((*t).to_string()).or_insert_with(|| algn.clone());
            }
        }
    }
    map
}

/// Parse master-level default East Asian line-break (eaLnBrk) per placeholder
/// type from each placeholder shape's lstStyle > lvl1pPr @eaLnBrk
/// (ECMA-376 §21.1.2.2.7). Mirrors parse_master_alignments. xsd:boolean.
pub(crate) fn parse_master_ea_ln_brk(root: roxmltree::Node<'_, '_>) -> HashMap<String, bool> {
    parse_master_ea_ln_brk_tier(root, true)
}

pub(crate) fn parse_master_ea_ln_brk_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> HashMap<String, bool> {
    let mut map = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(v) = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"))
            .and_then(|lp| attr(&lp, "eaLnBrk"))
        {
            map.entry(ph_type).or_insert(v == "1" || v == "true");
        }
    }
    inherit_master_placeholder_classes(&mut map);
    // txStyles fallback (the same class mapping as the other paragraph maps).
    for (style, types) in tx_style_nodes(root) {
        if let Some(v) = child(style, "lvl1pPr").and_then(|lp| attr(&lp, "eaLnBrk")) {
            for t in types {
                map.entry((*t).to_string())
                    .or_insert(v == "1" || v == "true");
            }
        }
    }
    map
}

/// Master default font alignment per placeholder type from each placeholder
/// shape's lstStyle > lvl1pPr @fontAlgn, then the class list styles
/// (ECMA-376 §21.1.2.2.7). Mirrors `parse_master_ea_ln_brk_tier`; the raw
/// token is kept so an explicit `base` still overrides a later tier.
pub(crate) fn parse_master_font_algn_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(v) = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"))
            .and_then(|lp| attr(&lp, "fontAlgn"))
        {
            map.entry(ph_type).or_insert(v.to_string());
        }
    }
    inherit_master_placeholder_classes(&mut map);
    for (style, types) in tx_style_nodes(root) {
        if let Some(v) = child(style, "lvl1pPr").and_then(|lp| attr(&lp, "fontAlgn")) {
            for t in types {
                map.entry((*t).to_string()).or_insert_with(|| v.to_string());
            }
        }
    }
    map
}

/// Per-list-level Latin faces from the master, keyed by placeholder type: the
/// master placeholder lstStyle wins per level over the class list style
/// (ECMA-376 §19.3.1.52, §21.1.2.3.7). Tokens resolve against this master's
/// theme, never the presentation-level theme: with a second master on its own
/// theme, `+mj-lt` / `+mn-lt` text rendered in that master's fonts (#1620).
pub(crate) fn parse_master_level_faces(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    default_text_style: Option<roxmltree::Node<'_, '_>>,
) -> HashMap<String, LevelFaces> {
    parse_master_level_faces_tier(root, theme, default_text_style, true)
}

pub(crate) fn parse_master_level_faces_tier(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    default_text_style: Option<roxmltree::Node<'_, '_>>,
    with_placeholders: bool,
) -> HashMap<String, LevelFaces> {
    let mut map: HashMap<String, LevelFaces> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(tx_body) = child(sp, "txBody") {
            let faces = extract_level_faces(tx_body, theme);
            if has_any_level_face(&faces) {
                map.entry(ph_type).or_insert(faces);
            }
        }
    }
    inherit_master_placeholder_classes(&mut map);
    let generic: Vec<(LevelFaces, &[&str])> = class_style_nodes(root, default_text_style)
        .into_iter()
        .map(|(style, types)| (read_level_faces(style, theme), types))
        .collect();
    for (faces, types) in generic {
        for t in types {
            let merged = merge_level_faces(map.get(*t).unwrap_or(&LevelFaces::default()), &faces);
            if has_any_level_face(&merged) {
                map.insert((*t).to_owned(), merged);
            }
        }
    }
    map
}

/// Per-list-level default font sizes from the master, keyed by ph_type. The
/// master placeholder lstStyle wins per level over the class list style; a
/// level neither names stays `None` (hard default 18 pt, not the level-1 size:
/// a body paragraph at lvl2 of a bodyStyle that defined only lvl1 rendered at
/// 18 pt in PowerPoint, issue #1620).
pub(crate) fn parse_master_level_font_sizes(
    root: roxmltree::Node<'_, '_>,
    default_text_style: Option<roxmltree::Node<'_, '_>>,
) -> HashMap<String, LevelFontSizes> {
    parse_master_level_font_sizes_tier(root, default_text_style, true)
}

pub(crate) fn parse_master_level_font_sizes_tier(
    root: roxmltree::Node<'_, '_>,
    default_text_style: Option<roxmltree::Node<'_, '_>>,
    with_placeholders: bool,
) -> HashMap<String, LevelFontSizes> {
    let mut map: HashMap<String, LevelFontSizes> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(tx_body) = child(sp, "txBody") {
            let sizes = extract_level_font_sizes(tx_body);
            if has_any_level_size(&sizes) {
                map.entry(ph_type).or_insert(sizes);
            }
        }
    }
    inherit_master_placeholder_classes(&mut map);
    let generic: Vec<(LevelFontSizes, &[&str])> = class_style_nodes(root, default_text_style)
        .into_iter()
        .map(|(style, types)| (read_level_font_sizes(style), types))
        .collect();
    for (sizes, types) in generic {
        for t in types {
            let merged = merge_level_sizes(map.get(*t).unwrap_or(&[None; 9]), &sizes);
            if has_any_level_size(&merged) {
                map.insert((*t).to_owned(), merged);
            }
        }
    }
    map
}

/// Per-list-level default run colours from the master placeholder cascade.
/// Per-placeholder list styles win per level; missing levels fall back to the
/// matching class list style.
pub(crate) fn parse_master_level_colors(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> HashMap<String, LevelColors> {
    parse_master_level_colors_tier(root, theme, true)
}

pub(crate) fn parse_master_level_colors_tier(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    with_placeholders: bool,
) -> HashMap<String, LevelColors> {
    let mut specific: HashMap<String, LevelColors> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(tx_body) = child(sp, "txBody") {
            let colors = extract_level_colors(tx_body, theme);
            if has_any_level_color(&colors) {
                specific.entry(ph_type).or_insert(colors);
            }
        }
    }

    let mut generic: HashMap<String, LevelColors> = HashMap::new();
    for (style_node, ph_types) in tx_style_nodes(root) {
        let colors = read_level_colors(style_node, theme);
        if has_any_level_color(&colors) {
            for ph_type in ph_types {
                generic.insert((*ph_type).to_owned(), colors.clone());
            }
        }
    }

    for (ph_type, colors) in generic {
        specific
            .entry(ph_type)
            .and_modify(|current| *current = merge_level_colors(current, &colors))
            .or_insert(colors);
    }
    specific
}

/// Per-level CT_TextCharacterProperties from the master placeholder lstStyle
/// and txStyles (`tx_style_nodes`). Merge each property so a partial master
/// placeholder defRPr still receives the omitted fields from its matching
/// txStyle (§21.1.2.4). The Latin face and size of these levels are replaced
/// by the resolved chain in shape.rs.
pub(crate) fn parse_master_level_run_properties(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_part: &str,
) -> MasterLevelRunProperties {
    parse_master_level_run_properties_tier(root, theme, master_rels, master_part, true)
}

pub(crate) fn parse_master_level_run_properties_tier(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_part: &str,
    with_placeholders: bool,
) -> MasterLevelRunProperties {
    let mut specific = HashMap::new();
    let mut styles: HashMap<String, LevelRunProperties> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(body) = child(sp, "txBody") {
            let props = extract_level_run_properties_with_rels(body, theme, master_rels)
                .map(|p| p.with_part_targets(master_part));
            if has_any_level_run_properties(&props) {
                specific.entry(ph_type).or_insert(props);
            }
        }
    }
    // Observed (#1620 controls): an obj and a typeless slot bound to a layout
    // slot took the master BODY placeholder's colour, bold, italic, underline
    // and caps, exactly like a body slot.
    inherit_master_placeholder_classes(&mut specific);
    for (style, ph_types) in tx_style_nodes(root) {
        let props = read_level_run_properties_with_rels(style, theme, master_rels)
            .map(|p| p.with_part_targets(master_part));
        if has_any_level_run_properties(&props) {
            for ph_type in ph_types {
                styles.insert((*ph_type).to_owned(), props.clone());
                specific
                    .entry((*ph_type).to_owned())
                    .and_modify(|own| *own = merge_level_run_properties(own, &props))
                    .or_insert_with(|| props.clone());
            }
        }
    }
    let share = |levels: HashMap<String, LevelRunProperties>| {
        levels
            .into_iter()
            .map(|(key, value)| (key, value.into()))
            .collect()
    };
    MasterLevelRunProperties {
        placeholders: share(specific),
        styles: share(styles),
    }
}

/// Master character properties per placeholder type, in two tiers.
/// `placeholders`: the master placeholder lstStyle merged over txStyles, for a
/// slide placeholder bound to a layout slot. `styles`: txStyles alone, for a
/// slide placeholder whose idx has NO layout slot. Observed (#1620 controls):
/// body, obj and typeless placeholders with an unmatched idx took bodyStyle's
/// colour, bold, italic, underline and caps and NOT those of the master body
/// placeholder's lstStyle, while the same types bound to a layout slot took
/// the master placeholder's values.
#[derive(Clone, Default, serde::Serialize)]
pub(crate) struct MasterLevelRunProperties {
    pub(crate) placeholders: HashMap<String, SharedLevelRunProperties>,
    pub(crate) styles: HashMap<String, SharedLevelRunProperties>,
}

/// Per-list-level paragraph indents (`marL`/`marR`/`indent`, EMU) from the
/// master, keyed by ph_type: master placeholder lstStyle, then the class list
/// style (ECMA-376 §21.1.2.4.13).
pub(crate) fn parse_master_level_indents(
    root: roxmltree::Node<'_, '_>,
) -> HashMap<String, LevelIndents> {
    parse_master_level_indents_tier(root, true)
}

pub(crate) fn parse_master_level_indents_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> HashMap<String, LevelIndents> {
    let mut map: HashMap<String, LevelIndents> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(tx_body) = child(sp, "txBody") {
            let indents = extract_level_indents(tx_body);
            if has_any_level_indent(&indents) {
                map.entry(ph_type).or_insert(indents);
            }
        }
    }
    inherit_master_placeholder_classes(&mut map);
    for (style_node, ph_types) in tx_style_nodes(root) {
        let indents = read_level_indents(style_node);
        if has_any_level_indent(&indents) {
            // Per level and axis: a master placeholder that indents only
            // some levels keeps the class style's other levels (§21.1.2.4).
            for ph_type in ph_types {
                map.entry(ph_type.to_string())
                    .and_modify(|existing| *existing = merge_level_indents(existing, &indents))
                    .or_insert(indents);
            }
        }
    }
    map
}

/// Per-list-level bullets from the master, keyed by ph_type: a master
/// placeholder's lstStyle bullets merged per group over the class list style
/// (ECMA-376 §19.7.10 / §21.1.2.4).
pub(crate) fn parse_master_level_bullets(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_part: &str,
    zip: &mut PptxZip,
) -> HashMap<String, LevelBullets> {
    parse_master_level_bullets_tier(root, theme, master_rels, master_part, zip, true)
}

pub(crate) fn parse_master_level_bullets_tier(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_part: &str,
    zip: &mut PptxZip,
    with_placeholders: bool,
) -> HashMap<String, LevelBullets> {
    let mut map: HashMap<String, LevelBullets> = HashMap::new();

    // A master-level `<a:buBlip>` embed resolves against the master's rels +
    // part directory (ECMA-376 §21.1.2.4.2), mirroring the master background.
    let mut resolve_blip = |rid: &str| -> Option<String> {
        let target = master_rels.get(rid)?;
        let path = resolve_path(master_part, target);
        // Existence check only (central directory, no inflate): a listed but
        // missing rId falls through to Bullet::Inherit.
        zip.index_for_name(&path)?;
        Some(path)
    };

    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(tx_body) = child(sp, "txBody") {
            let bullets = extract_level_bullets(tx_body, theme, &mut resolve_blip);
            if has_any_level_bullet(&bullets) {
                map.entry(ph_type).or_insert(bullets);
            }
        }
    }

    // An obj slot took the master body placeholder's buNone, not the
    // bodyStyle bullet (#1620 controls).
    inherit_master_placeholder_classes(&mut map);
    // txStyles fallback, resolved per bullet group: a per-shape entry that
    // declares only a marker inherits its colour/size/font from the matching
    // txStyles level (ECMA-376 §21.1.2.4, the four groups are independent).
    for (style_node, ph_types) in tx_style_nodes(root) {
        let bullets = read_level_bullets(style_node, theme, &mut resolve_blip);
        if has_any_level_bullet(&bullets) {
            for ph_type in ph_types {
                map.entry(ph_type.to_string())
                    .and_modify(|existing| *existing = merge_level_bullets(existing, &bullets))
                    .or_insert_with(|| bullets.clone());
            }
        }
    }
    map
}

/// Parse default bold/italic/caps/reflection from the class list styles
/// > lvl1pPr > defRPr. Keyed by ph_type; only explicit attributes populate.
type MasterTxStyleRunProperties = (
    HashMap<String, bool>,
    HashMap<String, bool>,
    HashMap<String, String>,
    HashMap<String, Reflection>,
);

pub(crate) fn parse_master_txstyle_run_properties(
    root: roxmltree::Node<'_, '_>,
) -> MasterTxStyleRunProperties {
    parse_master_txstyle_run_properties_tier(root, true)
}

pub(crate) fn parse_master_txstyle_run_properties_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> MasterTxStyleRunProperties {
    let mut bold_map: HashMap<String, bool> = HashMap::new();
    let mut italic_map: HashMap<String, bool> = HashMap::new();
    // ECMA-376 §21.1.2.3.9, ST_TextCapsType §20.1.10.64: cap="all"/"small"
    // on the class style defRPr — e.g. a template titleStyle with cap="all"
    // upper-cases every title.
    let mut caps_map: HashMap<String, String> = HashMap::new();
    let mut reflection_map: HashMap<String, Reflection> = HashMap::new();
    let mut read = |def_rpr: Option<roxmltree::Node<'_, '_>>, types: &[&str]| {
        let b = def_rpr
            .and_then(|rp| attr(&rp, "b"))
            .map(|v| v == "1" || v == "true");
        let i = def_rpr
            .and_then(|rp| attr(&rp, "i"))
            .map(|v| v == "1" || v == "true");
        let c = def_rpr
            .and_then(|rp| attr(&rp, "cap"))
            .filter(|v| v == "all" || v == "small");
        let reflection = def_rpr
            .and_then(|rp| child(rp, "effectLst"))
            .and_then(parse_reflection);
        for t in types {
            if let Some(bv) = b {
                bold_map.entry(t.to_string()).or_insert(bv);
            }
            if let Some(iv) = i {
                italic_map.entry(t.to_string()).or_insert(iv);
            }
            if let Some(cv) = &c {
                caps_map.entry(t.to_string()).or_insert(cv.clone());
            }
            if let Some(value) = &reflection {
                reflection_map
                    .entry(t.to_string())
                    .or_insert_with(|| value.clone());
            }
        }
    };
    // Master placeholder lstStyle first (with the class mapping), then txStyles.
    let placeholders = master_placeholder_shapes(root, with_placeholders);
    let lvl1_def_rpr = |sp| master_placeholder_lvl1(sp).and_then(|lp| child(lp, "defRPr"));
    for (ph_type, sp) in &placeholders {
        read(lvl1_def_rpr(*sp), &[ph_type.as_str()]);
    }
    for (source, targets) in MASTER_PLACEHOLDER_CLASSES {
        if let Some((_, sp)) = placeholders.iter().find(|(t, _)| t == source) {
            read(lvl1_def_rpr(*sp), targets);
        }
    }
    for (style_node, ph_types) in tx_style_nodes(root) {
        read(
            child(style_node, "lvl1pPr").and_then(|lp| child(lp, "defRPr")),
            ph_types,
        );
    }
    (bold_map, italic_map, caps_map, reflection_map)
}

/// Parse default text color from the master placeholder lstStyle > lvl1pPr >
/// defRPr, then the class list styles. Keyed by ph_type.
pub(crate) fn parse_master_txstyle_color(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> HashMap<String, String> {
    parse_master_txstyle_color_tier(root, theme, true)
}

pub(crate) fn parse_master_txstyle_color_tier(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    with_placeholders: bool,
) -> HashMap<String, String> {
    let mut map: HashMap<String, String> = HashMap::new();
    for (ph_type, sp) in master_placeholder_shapes(root, with_placeholders) {
        if let Some(color) = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"))
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| text_property_color(rp, theme))
        {
            map.entry(ph_type).or_insert(color);
        }
    }
    // Authored txStyles only: the built-in template styles of a master
    // without them colour their own list levels (`parse_master_level_colors`),
    // and a title's level-1 fill must not become the default of its deeper
    // levels, which keep the hard default.
    let authored_tx_styles = child(root, "txStyles").is_some();
    for (style_node, ph_types) in tx_style_nodes(root).into_iter().filter(|_| authored_tx_styles) {
        if let Some(color) = child(style_node, "lvl1pPr")
            .and_then(|lp| child(lp, "defRPr"))
            .and_then(|rp| text_property_color(rp, theme))
        {
            for ph_type in ph_types {
                map.entry(ph_type.to_string()).or_insert(color.clone());
            }
        }
    }
    map
}

/// Per-level paragraph spacing (spcBef / spcAft / lnSpc) from the master,
/// keyed by ph_type: a master placeholder's lstStyle (with the class mapping
/// of `MASTER_PLACEHOLDER_CLASSES`) wins per level and property over the
/// txStyles class style.
pub(crate) fn parse_master_txstyle_spacing(
    root: roxmltree::Node<'_, '_>,
) -> HashMap<String, LevelSpacing> {
    parse_master_txstyle_spacing_tier(root, true)
}

pub(crate) fn parse_master_txstyle_spacing_tier(
    root: roxmltree::Node<'_, '_>,
    with_placeholders: bool,
) -> HashMap<String, LevelSpacing> {
    let mut map: HashMap<String, LevelSpacing> = HashMap::new();
    let mut read = |list_style: Option<roxmltree::Node<'_, '_>>, types: &[&str]| {
        let Some(spacing) = list_style.map(LevelSpacing::read).filter(|s| !s.is_empty()) else {
            return;
        };
        for ph_type in types {
            let entry = map.entry((*ph_type).to_owned()).or_default();
            *entry = entry.or(&spacing);
        }
    };
    // Master placeholder lstStyle first (with the class mapping), then txStyles.
    let placeholders = master_placeholder_shapes(root, with_placeholders);
    for (ph_type, sp) in &placeholders {
        read(placeholder_list_style(*sp), &[ph_type.as_str()]);
    }
    for (source, targets) in MASTER_PLACEHOLDER_CLASSES {
        if let Some((_, sp)) = placeholders.iter().find(|(t, _)| t == source) {
            read(placeholder_list_style(*sp), targets);
        }
    }
    for (style_node, ph_types) in tx_style_nodes(root) {
        read(Some(style_node), ph_types);
    }
    map
}

/// The txStyles-only tier of every list-style map, for a slide placeholder
/// whose idx has no layout slot. Observed (#1620 controls): body, obj and
/// typeless placeholders with an unmatched idx took bodyStyle's colour, bold,
/// italic, underline and caps, not the master placeholder lstStyle's, while
/// the same types bound to a layout slot took the master placeholder's values.
/// Every per-level and paragraph map follows the same two tiers:
///
/// * bound to a layout slot: layout slot -> master placeholder (obj, subTitle
///   and typeless use the master body placeholder) -> class style;
/// * unmatched idx: class style only (this tier);
/// * dt/ftr/sldNum: face and size from defaultTextStyle, every other property
///   from otherStyle (`class_style_nodes` / `tx_style_nodes`).
#[derive(Clone, Default, serde::Serialize)]
pub(crate) struct MasterStyleTier {
    pub(crate) faces: HashMap<String, LevelFaces>,
    pub(crate) sizes: HashMap<String, LevelFontSizes>,
    pub(crate) colors: HashMap<String, LevelColors>,
    pub(crate) indents: HashMap<String, LevelIndents>,
    pub(crate) bullets: HashMap<String, LevelBullets>,
    pub(crate) alignment: HashMap<String, String>,
    pub(crate) ea_ln_brk: HashMap<String, bool>,
    /// Class-tier fontAlgn (txStyles only), like `ea_ln_brk`.
    pub(crate) font_algn: HashMap<String, String>,
    /// fontAlgn from the master placeholders' lstStyle, then txStyles: the
    /// master fallback for a layout placeholder (like `master_ea_ln_brk`).
    pub(crate) placeholder_font_algn: HashMap<String, String>,
    pub(crate) spacing: HashMap<String, LevelSpacing>,
    pub(crate) color: HashMap<String, String>,
    pub(crate) bold: HashMap<String, bool>,
    pub(crate) italic: HashMap<String, bool>,
    pub(crate) caps: HashMap<String, String>,
    pub(crate) reflection: HashMap<String, Reflection>,
}

impl MasterStyleTier {
    /// Resolve against `theme` (the master's, or a clrMapOvr slide's).
    pub(crate) fn parse(
        root: roxmltree::Node<'_, '_>,
        theme: &HashMap<String, String>,
        master_rels: &HashMap<String, String>,
        master_part: &str,
        default_text_style: Option<roxmltree::Node<'_, '_>>,
        zip: &mut PptxZip,
    ) -> Self {
        let (bold, italic, caps, reflection) =
            parse_master_txstyle_run_properties_tier(root, false);
        MasterStyleTier {
            faces: parse_master_level_faces_tier(root, theme, default_text_style, false),
            sizes: parse_master_level_font_sizes_tier(root, default_text_style, false),
            colors: parse_master_level_colors_tier(root, theme, false),
            indents: parse_master_level_indents_tier(root, false),
            bullets: parse_master_level_bullets_tier(
                root,
                theme,
                master_rels,
                master_part,
                zip,
                false,
            ),
            alignment: parse_master_alignments_tier(root, false),
            ea_ln_brk: parse_master_ea_ln_brk_tier(root, false),
            font_algn: parse_master_font_algn_tier(root, false),
            placeholder_font_algn: parse_master_font_algn_tier(root, true),
            spacing: parse_master_txstyle_spacing_tier(root, false),
            color: parse_master_txstyle_color_tier(root, theme, false),
            bold,
            italic,
            caps,
            reflection,
        }
    }

    pub(crate) fn get<'a, T>(map: &'a HashMap<String, T>, ph_type: &str) -> Option<&'a T> {
        map.get(ph_type)
            .or_else(|| if ph_type == "obj" { map.get("") } else { None })
    }
}

/// The presentation `defaultTextStyle` levels that ordinary (non-placeholder)
/// text inherits: text boxes, autoshapes and table cells on slides, layouts and
/// masters (ECMA-376 §19.2.1.8). Tokens resolve against the master theme.
/// Observed (issue #1620): a text box took the level's face and size; a level
/// without `<a:latin>` rendered in Arial at the level's size, not the level-1
/// face and not the theme minor font.
#[derive(Debug, Clone, Default, serde::Serialize)]
pub(crate) struct DefaultTextLevels {
    pub(crate) faces: LevelFaces,
    pub(crate) sizes: LevelFontSizes,
    /// Authored `<a:ea>` / `<a:cs>` typefaces per level (tokens stay tokens;
    /// they resolve per run language, issue #1627).
    pub(crate) east_asian: LevelFaces,
    pub(crate) complex_script: LevelFaces,
    /// Per-level `lang` / `altLang` of the level defRPr.
    pub(crate) lang: LevelFaces,
    pub(crate) alt_lang: LevelFaces,
    /// The marL a plain paragraph takes when nothing in its own cascade sets
    /// one: the level's marL, else 0 (issue #1628 controls: text boxes whose
    /// defaultTextStyle level set no marL started at the inset at levels 2
    /// and 3).
    pub(crate) mar_l: [i64; 9],
    /// The text colour per level, complete at every level
    /// (`default_text_level_colors`). Empty only outside a presentation
    /// (standalone shapes, SmartArt drawings), where no defaultTextStyle applies.
    pub(crate) colors: LevelColors,
    /// The list style table-cell text inherits (`parse_table_text_levels`).
    pub(crate) table: TableTextLevels,
}

/// The text colour each defaultTextStyle level gives ordinary text, resolved
/// against `theme`, the slide's effective colour map (§19.3.1.6 clrMap, then
/// the layout's and slide's §19.3.1.7 clrMapOvr).
///
/// Observed (issue #1989, PowerPoint PDF export of synthetic decks whose
/// master swaps the map to bg1=dk1 tx1=lt1 bg2=dk2 tx2=lt2, with distinctive
/// dk1/lt1/dk2/lt2 values and every source in a distinct colour):
/// * Text boxes, autoshapes without a fontRef colour, grouped text boxes and
///   the non-placeholder text of layouts and masters took their level's fill
///   through the effective map: lvl1 tx1 rendered lt1, lvl2 tx2 rendered lt2;
///   a slide or layout overrideClrMapping flipped them back to dk1/dk2.
/// * A level without a fill, under a defaultTextStyle that exists, rendered
///   000000 at every level and under both maps; not dk1 (a 1F3F5F dk1), not
///   tx1, not the defPPr fill and not the master otherStyle fill.
/// * A presentation without a defaultTextStyle rendered tx1 at every level.
/// * The master otherStyle fills never applied to this text.
pub(crate) fn default_text_level_colors(
    default_text_style: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> LevelColors {
    match default_text_style {
        Some(node) => read_level_colors(node, theme)
            .map(|color| Some(color.unwrap_or_else(|| HARD_DEFAULT_TEXT_COLOR.to_owned()))),
        None => {
            let tx1 = theme
                .get("tx1")
                .cloned()
                .unwrap_or_else(|| HARD_DEFAULT_TEXT_COLOR.to_owned());
            std::array::from_fn(|_| Some(tx1.clone()))
        }
    }
}

/// The list-level properties table-cell text inherits: the slide master's
/// `otherStyle`, which ECMA-376 §19.3.1.35 describes as the style of text
/// that is neither title nor body.
///
/// Observed (issue #1628, PowerPoint's reference PDF export), with every
/// source at a distinct size:
/// * A cell paragraph took the size, alignment and marL of its own level of
///   the master's otherStyle, in every table style (built-in, custom, none)
///   and in placeholder graphic frames. The size did not come from
///   defaultTextStyle, bodyStyle or the layout slot.
/// * A level otherStyle does not define ended at 18 pt with no indent. Its
///   defPPr had no effect.
/// * otherStyle's Latin face and colour did not apply: the face is the table
///   style's, else the theme minor font (`complete_table_cell_faces`).
/// * A master with no otherStyle (no txStyles, or txStyles without one) laid
///   cells out at 18 pt with marL 0.5" per level, PowerPoint's default
///   otherStyle.
///
/// Only size, alignment and indents were observable; the other paragraph and
/// character properties of otherStyle are not applied here.
#[derive(Debug, Clone, serde::Serialize)]
pub(crate) struct TableTextLevels {
    pub(crate) sizes: LevelFontSizes,
    pub(crate) indents: LevelIndents,
    pub(crate) alignments: LevelAlignments,
}

impl Default for TableTextLevels {
    /// PowerPoint's default otherStyle: 18 pt, marL 0.5" per level.
    fn default() -> Self {
        let mut indents: LevelIndents = Default::default();
        for (lvl, indent) in indents.iter_mut().enumerate() {
            indent.mar_l = Some(lvl as i64 * 457_200);
            indent.indent = Some(0);
        }
        TableTextLevels {
            sizes: [Some(HARD_DEFAULT_FONT_SIZE); 9],
            indents,
            alignments: Default::default(),
        }
    }
}

/// Build [`TableTextLevels`] from a slide master's `txStyles/otherStyle`.
pub(crate) fn parse_table_text_levels(root: roxmltree::Node<'_, '_>) -> TableTextLevels {
    let Some(other) = child(root, "txStyles").and_then(|tx| child(tx, "otherStyle")) else {
        return TableTextLevels::default();
    };
    let indents = read_level_indents(other);
    TableTextLevels {
        sizes: complete_level_sizes(&read_level_font_sizes(other)),
        indents: std::array::from_fn(|lvl| LevelIndent {
            mar_l: Some(indents[lvl].mar_l.unwrap_or(0)),
            mar_r: indents[lvl].mar_r,
            indent: Some(indents[lvl].indent.unwrap_or(0)),
        }),
        alignments: read_level_alignments(other),
    }
}

/// Build [`DefaultTextLevels`]. When the presentation has no defaultTextStyle
/// at all PowerPoint behaves as if every level named the theme minor Latin
/// font at 18 pt with marL 0.5" per level: text boxes at lvl1 and lvl2 rendered in the master's minor
/// font at 18 pt (issue #1620), unlike a present level that omits the face.
pub(crate) fn parse_default_text_levels(
    default_text_style: Option<roxmltree::Node<'_, '_>>,
    theme: &HashMap<String, String>,
) -> DefaultTextLevels {
    match default_text_style {
        Some(node) => {
            let indents = read_level_indents(node);
            DefaultTextLevels {
                faces: read_level_faces(node, theme),
                sizes: read_level_font_sizes(node),
                mar_l: std::array::from_fn(|level| indents[level].mar_l.unwrap_or(0)),
                east_asian: read_level_defrpr_values(node, Some("ea"), "typeface"),
                complex_script: read_level_defrpr_values(node, Some("cs"), "typeface"),
                lang: read_level_defrpr_values(node, None, "lang"),
                alt_lang: read_level_defrpr_values(node, None, "altLang"),
                colors: default_text_level_colors(Some(node), theme),
                table: TableTextLevels::default(),
            }
        }
        None => DefaultTextLevels {
            faces: std::array::from_fn(|_| resolve_latin_face("+mn-lt", theme)),
            sizes: [Some(HARD_DEFAULT_FONT_SIZE); 9],
            mar_l: DEFAULT_TEXT_STYLE_MAR_L,
            colors: default_text_level_colors(None, theme),
            table: TableTextLevels::default(),
            ..Default::default()
        },
    }
}

pub(crate) fn parse_master_transforms(root: roxmltree::Node<'_, '_>) -> HashMap<String, Transform> {
    let mut map = HashMap::new();
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            let ph_node = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph");
            if let Some(ph) = ph_node {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(xfrm) = child(sp, "spPr").and_then(|p| child(p, "xfrm")) {
                    map.entry(ph_type).or_insert_with(|| parse_xfrm(xfrm));
                }
            }
        }
    }
    map
}

// Seeds layout placeholders from the master's per-type defaults (transforms,
// alignment, spacing) before overlaying the layout's own placeholder props; the
// many maps are the master inheritance sources, threaded through as-is.
//
// Takes the already-parsed layout root (`<p:sldLayout>`) so `parse_layout` can
// parse the layout XML once and share the `Document` with the background +
// showMasterSp extractions (D4).
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_layout_placeholders(
    root: roxmltree::Node<'_, '_>,
    master_level_faces: &HashMap<String, LevelFaces>,
    default_text: &DefaultTextLevels,
    master_styles: &MasterStyleTier,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_colors: &HashMap<String, LevelColors>,
    master_level_run_properties: &MasterLevelRunProperties,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_body_pr: &HashMap<String, InheritedBodyPr>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_spacing: &HashMap<String, LevelSpacing>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    layout_part: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> LayoutPlaceholders {
    let theme = theme_source.colors();
    let mut lph = LayoutPlaceholders {
        master_by_type: master_transforms.clone(),
        by_type_master_level_faces: master_styles.faces.clone(),
        default_text: default_text.clone(),
        styles: master_styles.clone(),
        by_type_master_level_sizes: master_styles.sizes.clone(),
        by_type_master_level_colors: master_styles.colors.clone(),
        by_type_master_level_run_properties: master_level_run_properties.styles.clone(),
        by_type_master_level_indents: master_styles.indents.clone(),
        by_type_master_level_bullets: master_styles.bullets.clone(),
        by_type_master_anchor: master_anchors.clone(),
        by_type_master_body_pr: master_body_pr.clone(),
        by_type_master_alignment: master_alignments.clone(),
        by_type_master_ea_ln_brk: master_ea_ln_brk.clone(),
        by_type_master_spacing: master_spacing.clone(),
        ..Default::default()
    };

    let sp_tree = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "spTree");
    let sp_tree = match sp_tree {
        Some(n) => n,
        None => return lph,
    };

    let mut seen_types: std::collections::HashSet<String> = std::collections::HashSet::new();
    for sp in sp_tree
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "sp")
    {
        let ph_node = sp
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "ph");
        let layout_ph_type = ph_node
            .and_then(|ph| attr(&ph, "type"))
            .unwrap_or_else(|| "obj".to_owned());
        let sp_pr = match child(sp, "spPr") {
            Some(n) => n,
            None => continue,
        };
        // xfrm may be absent (placeholder inherits transform from master); parse if present
        let t_opt: Option<Transform> = child(sp_pr, "xfrm").map(parse_xfrm);

        // Extract layout-level defaults from the placeholder's txBody > lstStyle > lvl1pPr
        let layout_lvl1_ppr: Option<roxmltree::Node<'_, '_>> = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .and_then(|ls| child(ls, "lvl1pPr"));
        let layout_def_rpr: Option<roxmltree::Node<'_, '_>> =
            layout_lvl1_ppr.and_then(|lp| child(lp, "defRPr"));
        let layout_level_faces: LevelFaces = child(sp, "txBody")
            .map(|tx_body| extract_level_faces(tx_body, theme))
            .unwrap_or_default();
        let has_text_body = child(sp, "txBody").is_some();
        // Per-level sizes from the layout placeholder's own lstStyle (all
        // lvlNpPr), used to give nested bullets their shrinking sizes.
        let layout_level_sizes: LevelFontSizes = child(sp, "txBody")
            .map(extract_level_font_sizes)
            .unwrap_or([None; 9]);
        let layout_level_colors: LevelColors = child(sp, "txBody")
            .map(|tx_body| extract_level_colors(tx_body, theme))
            .unwrap_or_else(|| std::array::from_fn(|_| None));
        let layout_level_run_properties = child(sp, "txBody")
            .map(|tx_body| {
                extract_level_run_properties_with_rels(tx_body, theme, layout_rels)
                    .map(|p| p.with_part_targets(layout_part))
            })
            .unwrap_or_else(|| std::array::from_fn(|_| Default::default()));
        // Per-level indents (marL/marR/indent) from the layout placeholder's own
        // lstStyle, the inherited list-indent cascade (ECMA-376 §21.1.2.4.13).
        let layout_level_indents: LevelIndents = child(sp, "txBody")
            .map(extract_level_indents)
            .unwrap_or_default();
        // Per-level bullets from the layout placeholder's own lstStyle. A
        // level's `<a:buBlip>` embed (§21.1.2.4.2) resolves against the layout's
        // rels + part directory, mirroring the layout-spPr blipFill above.
        let mut resolve_layout_blip = |rid: &str| -> Option<String> {
            let target = layout_rels.get(rid)?;
            let path = resolve_path(layout_part, target);
            // Verify the part exists so a listed-but-missing rId yields None and
            // the bullet falls through to Bullet::Inherit (matches the variant's
            // doc comment), mirroring the master/layout background resolvers.
            // `index_for_name` reads the central directory only (no inflate),
            // unlike the former `read_zip_bytes` which decompressed and discarded.
            zip.index_for_name(&path)?;
            Some(path)
        };
        let layout_level_bullets: LevelBullets = child(sp, "txBody")
            .map(|tb| extract_level_bullets(tb, theme, &mut resolve_layout_blip))
            .unwrap_or_else(empty_level_bullets);
        let layout_bold = layout_def_rpr
            .and_then(|rp| attr(&rp, "b"))
            .map(|v| v == "1" || v == "true");
        let layout_italic = layout_def_rpr
            .and_then(|rp| attr(&rp, "i"))
            .map(|v| v == "1" || v == "true");
        let layout_caps = layout_def_rpr
            .and_then(|rp| attr(&rp, "cap"))
            .filter(|v| v == "all" || v == "small");
        let layout_reflection = layout_def_rpr
            .and_then(|rp| child(rp, "effectLst"))
            .and_then(parse_reflection);
        let layout_color: Option<String> =
            layout_def_rpr.and_then(|rp| text_property_color(rp, theme));
        let layout_alignment: Option<String> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "algn"))
            .map(|a| a.to_string());
        // ECMA-376 §21.1.2.2.7 eaLnBrk from the layout placeholder's lvl1pPr.
        let layout_ea_ln_brk: Option<bool> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "eaLnBrk"))
            .map(|v| v == "1" || v == "true");
        let layout_font_algn: Option<String> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "fontAlgn"))
            .map(|v| v.to_string());
        let layout_spacing = child(sp, "txBody")
            .and_then(|tb| child(tb, "lstStyle"))
            .map(LevelSpacing::read)
            .filter(|spacing| !spacing.is_empty());

        let layout_body_pr = child(sp, "txBody").and_then(|tb| child(tb, "bodyPr"));
        // Layout bodyPr anchor; fall back to master anchor map.
        let layout_anchor: Option<String> = layout_body_pr
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a.to_string());
        // Every modelled bodyPr value merges layout → master attribute by
        // attribute (see InheritedBodyPr for the PowerPoint evidence).
        let effective_body_pr = layout_body_pr
            .map(InheritedBodyPr::from_body_pr)
            .unwrap_or_default()
            .or(&master_body_pr
                .get(&layout_ph_type)
                .cloned()
                .unwrap_or_default());

        // A picture placeholder inherits the same CT_ShapeProperties component
        // cascade as an ordinary picture. Resolve the layout's local/style
        // tiers once, then retain that bundle alongside its blipFill.
        let layout_picture_properties =
            resolve_picture_shape_properties(Some(sp_pr), child(sp, "style"), None, theme_source);
        let layout_stroke = layout_picture_properties.stroke.clone();

        // Layout spPr fill (solidFill / noFill / gradFill / pattFill). The
        // slide-level placeholder shape inherits this when its own `<p:spPr>` is
        // empty — that's how a "tinted body placeholder" carries through to the
        // slide. We deliberately exclude grpFill here (group inheritance is
        // resolved at slide parse time, not from the layout).
        let layout_fill: Option<Fill> = parse_fill(sp_pr, theme);
        let layout_xfrm = child(sp_pr, "xfrm").map(parse_xfrm).unwrap_or_default();
        let layout_geometry =
            InheritedShapeGeometry::from_sp_pr(sp_pr, layout_xfrm.cx as f64, layout_xfrm.cy as f64);

        // Layout spPr > blipFill → image that bleeds through when the slide's
        // matching placeholder has no own blipFill (picture placeholder inheritance).
        let layout_blip_fill: Option<InheritedBlipFill> = child(sp_pr, "blipFill").and_then(|bf| {
            let rid = child(bf, "blip").and_then(|b| attr_r(&b, "embed"))?;
            let rel_target = layout_rels.get(&rid)?;
            let image_path = resolve_path(layout_part, rel_target);
            // Verify the part exists so a dangling rId yields None (no inherited
            // fill), preserving the prior data-URL behaviour. `index_for_name`
            // reads the central directory only (no inflate), unlike the former
            // `read_zip_bytes` which decompressed the entry just to discard it.
            zip.index_for_name(&image_path)?;
            let mime_type = mime_from_ext(&image_path).to_owned();
            Some(InheritedBlipFill {
                image_path,
                mime_type,
                src_rect: parse_src_rect(bf),
                alpha: None, /* alpha is baked by the ordered blip chain */
                // §20.1.8.23 duotone on the layout placeholder blipFill, resolved
                // through the theme; inherited onto the slide picture placeholder.
                duotone: parse_blip_duotone(
                    bf,
                    &PptxSchemeResolver { theme },
                    ooxml_common::color::TintMode::PowerPointLinear,
                ),
                // CT_Blip pixel effects travel with the inherited blipFill.
                blip_effects: {
                    let mut effects = parse_blip_effects(
                        bf,
                        &PptxSchemeResolver { theme },
                        ooxml_common::color::TintMode::PowerPointLinear,
                    );
                    ooxml_common::blip::resolve_blip_effect_images(&mut effects, &mut |rid| {
                        layout_rels
                            .get(rid)
                            .map(|target| resolve_path(layout_part, target))
                    });
                    effects
                },
            })
        });

        if let Some(ph) = ph_node {
            // CT_Placeholder defaults an omitted @type to `obj`. The idx binds
            // the slide placeholder to this layout slot, but does not permit a
            // same-numbered master placeholder of another type to rewrite the
            // schema value. Real PowerPoint layouts can reuse an idx for a
            // content slot where the master uses it for date/footer metadata.
            let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());
            let ph_type = attr(&ph, "type").unwrap_or_else(|| "obj".to_owned());

            // The type-keyed maps below keep the first slot of each type.
            let first_of_type = seen_types.insert(ph_type.clone());
            if !has_text_body && first_of_type {
                lph.by_type_without_text_body.insert(ph_type.clone());
            }
            if let Some(idx) = ph_idx {
                let new_slot = !lph.by_idx_placeholder_type.contains_key(&idx);
                lph.by_idx_placeholder_type
                    .entry(idx)
                    .or_insert_with(|| ph_type.clone());
                if new_slot && !has_text_body {
                    lph.by_idx_without_text_body.insert(idx);
                }
                if let Some(ref t) = t_opt {
                    lph.by_idx.entry(idx).or_insert_with(|| t.clone());
                }
                let level_faces = merge_level_faces(
                    &layout_level_faces,
                    master_level_faces
                        .get(&ph_type)
                        .unwrap_or(&LevelFaces::default()),
                );
                if has_any_level_face(&level_faces) {
                    lph.by_idx_level_faces.entry(idx).or_insert(level_faces);
                }
                // Per-level: layout lstStyle wins per level, else master.
                let level_sizes = merge_level_sizes(
                    &layout_level_sizes,
                    master_level_font_sizes.get(&ph_type).unwrap_or(&[None; 9]),
                );
                if has_any_level_size(&level_sizes) {
                    lph.by_idx_level_sizes.entry(idx).or_insert(level_sizes);
                }
                let empty_colors: LevelColors = std::array::from_fn(|_| None);
                let level_colors = merge_level_colors(
                    &layout_level_colors,
                    master_level_colors.get(&ph_type).unwrap_or(&empty_colors),
                );
                if has_any_level_color(&level_colors) {
                    lph.by_idx_level_colors.entry(idx).or_insert(level_colors);
                }
                let empty_run = std::array::from_fn(|_| Default::default());
                let level_run = merge_level_run_properties(
                    &layout_level_run_properties,
                    master_level_run_properties
                        .placeholders
                        .get(&ph_type)
                        .map(|levels| &**levels)
                        .unwrap_or(&empty_run),
                );
                if has_any_level_run_properties(&level_run) {
                    lph.by_idx_level_run_properties
                        .entry(idx)
                        .or_insert(level_run.into());
                }
                // Per-level indents: layout lstStyle wins per axis/level, else master.
                let level_indents = merge_level_indents(
                    &layout_level_indents,
                    master_level_indents
                        .get(&ph_type)
                        .unwrap_or(&Default::default()),
                );
                if has_any_level_indent(&level_indents) {
                    lph.by_idx_level_indents.entry(idx).or_insert(level_indents);
                }
                // Per-level bullets: layout lstStyle wins per level, else master.
                let empty_bul = empty_level_bullets();
                let level_bullets = merge_level_bullets(
                    &layout_level_bullets,
                    master_level_bullets.get(&ph_type).unwrap_or(&empty_bul),
                );
                if has_any_level_bullet(&level_bullets) {
                    lph.by_idx_level_bullets.entry(idx).or_insert(level_bullets);
                }
                if let Some(ref s) = layout_stroke {
                    lph.by_idx_stroke.entry(idx).or_insert(s.clone());
                }
                if !layout_picture_properties.is_empty() {
                    lph.by_idx_picture_properties
                        .entry(idx)
                        .or_insert_with(|| layout_picture_properties.clone());
                }
                if let Some(v) = layout_spacing.clone() {
                    lph.by_idx_spacing.entry(idx).or_insert(v);
                }
                if !effective_body_pr.is_empty() {
                    lph.by_idx_body_pr
                        .entry(idx)
                        .or_insert_with(|| effective_body_pr.clone());
                }
                if let Some(ref bf) = layout_blip_fill {
                    lph.by_idx_blip_fill.entry(idx).or_insert(bf.clone());
                }
                if let Some(ref c) = layout_color {
                    lph.by_idx_color.entry(idx).or_insert(c.clone());
                }
                if let Some(ref f) = layout_fill {
                    lph.by_idx_fill.entry(idx).or_insert(f.clone());
                }
                if let Some(ref geometry) = layout_geometry {
                    lph.by_idx_geometry
                        .entry(idx)
                        .or_insert_with(|| geometry.clone());
                }
                // Alignment for this idx: layout's own algn, else master per-type
                // (incl. master txStyles, now folded into master_alignments).
                let idx_algn = layout_alignment
                    .clone()
                    .or_else(|| master_alignments.get(&ph_type).cloned());
                if let Some(a) = idx_algn {
                    lph.by_idx_alignment.entry(idx).or_insert(a);
                }
                let idx_font_algn = layout_font_algn.clone().or_else(|| {
                    MasterStyleTier::get(&master_styles.placeholder_font_algn, &ph_type).cloned()
                });
                if let Some(f) = idx_font_algn {
                    lph.by_idx_font_algn.entry(idx).or_insert(f);
                }
                // ECMA-376 §19.3.1.36: idx binds the slide placeholder to this
                // exact layout slot. Preserve its vertical anchor independently
                // from same-type siblings; fall back to the master for this type.
                let idx_anchor = layout_anchor
                    .clone()
                    .or_else(|| master_anchors.get(&ph_type).cloned());
                if let Some(a) = idx_anchor {
                    lph.by_idx_anchor.entry(idx).or_insert(a);
                }
            }
            let type_level_faces = merge_level_faces(
                &layout_level_faces,
                master_level_faces
                    .get(&ph_type)
                    .unwrap_or(&LevelFaces::default()),
            );
            if has_any_level_face(&type_level_faces) {
                lph.by_type_level_faces
                    .entry(ph_type.clone())
                    .or_insert(type_level_faces);
            }
            let type_level_sizes = merge_level_sizes(
                &layout_level_sizes,
                master_level_font_sizes.get(&ph_type).unwrap_or(&[None; 9]),
            );
            if has_any_level_size(&type_level_sizes) {
                lph.by_type_level_sizes
                    .entry(ph_type.clone())
                    .or_insert(type_level_sizes);
            }
            let empty_colors: LevelColors = std::array::from_fn(|_| None);
            let type_level_colors = merge_level_colors(
                &layout_level_colors,
                master_level_colors.get(&ph_type).unwrap_or(&empty_colors),
            );
            if has_any_level_color(&type_level_colors) {
                lph.by_type_level_colors
                    .entry(ph_type.clone())
                    .or_insert(type_level_colors);
            }
            let empty_run = std::array::from_fn(|_| Default::default());
            let type_level_run = merge_level_run_properties(
                &layout_level_run_properties,
                master_level_run_properties
                    .placeholders
                    .get(&ph_type)
                    .map(|levels| &**levels)
                    .unwrap_or(&empty_run),
            );
            if has_any_level_run_properties(&type_level_run) {
                lph.by_type_level_run_properties
                    .entry(ph_type.clone())
                    .or_insert(type_level_run.into());
            }
            let type_level_indents = merge_level_indents(
                &layout_level_indents,
                master_level_indents
                    .get(&ph_type)
                    .unwrap_or(&Default::default()),
            );
            if has_any_level_indent(&type_level_indents) {
                lph.by_type_level_indents
                    .entry(ph_type.clone())
                    .or_insert(type_level_indents);
            }
            let empty_bul_t = empty_level_bullets();
            let type_level_bullets = merge_level_bullets(
                &layout_level_bullets,
                master_level_bullets.get(&ph_type).unwrap_or(&empty_bul_t),
            );
            if has_any_level_bullet(&type_level_bullets) {
                lph.by_type_level_bullets
                    .entry(ph_type.clone())
                    .or_insert(type_level_bullets);
            }
            if let Some(b) = layout_bold {
                lph.by_type_bold.entry(ph_type.clone()).or_insert(b);
            }
            if let Some(i) = layout_italic {
                lph.by_type_italic.entry(ph_type.clone()).or_insert(i);
            }
            if let Some(c) = layout_caps.clone() {
                lph.by_type_caps.entry(ph_type.clone()).or_insert(c);
            }
            if let Some(reflection) = layout_reflection.clone() {
                lph.by_type_reflection
                    .entry(ph_type.clone())
                    .or_insert(reflection);
            }
            if let Some(a) = layout_alignment {
                lph.by_type_alignment.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(e) = layout_ea_ln_brk {
                lph.by_type_ea_ln_brk.entry(ph_type.clone()).or_insert(e);
            }
            if let Some(f) = layout_font_algn.clone() {
                lph.by_type_font_algn.entry(ph_type.clone()).or_insert(f);
            }
            if let Some(v) = layout_spacing {
                // Per level and property, the first same-type slot that sets it.
                let entry = lph.by_type_spacing.entry(ph_type.clone()).or_default();
                *entry = entry.or(&v);
            }
            if !effective_body_pr.is_empty() {
                // Per field, the first same-type layout placeholder that sets it.
                let merged = match lph.by_type_body_pr.remove(&ph_type) {
                    Some(first) => first.or(&effective_body_pr),
                    None => effective_body_pr,
                };
                lph.by_type_body_pr.insert(ph_type.clone(), merged);
            }
            // Anchor: layout bodyPr > fall back to master anchor map
            let effective_anchor = layout_anchor
                .clone()
                .or_else(|| master_anchors.get(&ph_type).cloned());
            if let Some(a) = effective_anchor {
                lph.by_type_anchor.entry(ph_type.clone()).or_insert(a);
            }
            if let Some(s) = layout_stroke {
                lph.by_type_stroke.entry(ph_type.clone()).or_insert(s);
            }
            if !layout_picture_properties.is_empty() {
                lph.by_type_picture_properties
                    .entry(ph_type.clone())
                    .or_insert(layout_picture_properties);
            }
            if let Some(bf) = layout_blip_fill {
                lph.by_type_blip_fill.entry(ph_type.clone()).or_insert(bf);
            }
            if let Some(c) = layout_color {
                lph.by_type_color.entry(ph_type.clone()).or_insert(c);
            }
            if let Some(f) = layout_fill {
                lph.by_type_fill.entry(ph_type.clone()).or_insert(f);
            }
            if let Some(geometry) = layout_geometry {
                lph.by_type_geometry
                    .entry(ph_type.clone())
                    .or_insert(geometry);
            }
            if let Some(t) = t_opt {
                lph.by_type.entry(ph_type).or_insert(t);
            }
        }
    }

    // A slide placeholder can be intentionally unbound to a layout slot (for
    // example PowerPoint's idx=2^32-1 sentinel on a blank layout). In that case
    // it still inherits the matching master txStyles / placeholder defaults by
    // type. The loop above only materializes type entries that also occur in
    // the layout, so fill the absent types from the master after all layout
    // overlays have won.
    for (ph_type, value) in master_level_faces {
        lph.by_type_level_faces
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    for (ph_type, value) in master_level_font_sizes {
        lph.by_type_level_sizes
            .entry(ph_type.clone())
            .or_insert(*value);
    }
    for (ph_type, value) in master_level_colors {
        lph.by_type_level_colors
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    for (ph_type, value) in &master_level_run_properties.placeholders {
        lph.by_type_level_run_properties
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    for (ph_type, value) in master_level_indents {
        lph.by_type_level_indents
            .entry(ph_type.clone())
            .or_insert(*value);
    }
    for (ph_type, value) in master_level_bullets {
        lph.by_type_level_bullets
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    for (ph_type, value) in master_anchors {
        lph.by_type_anchor
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    lph
}

/// The layout XML parsed ONCE into the owned data a slide needs from its layout
/// (D4). Groups the three former per-slide layout re-parses in `parse_slide`:
/// placeholder inheritance (§19.3.1.39), the layout-level `<p:bg>` background,
/// and the layout's `showMasterSp` flag (§19.3.1.39). Holds no `roxmltree` node
/// (owned only), so it can be cached across slides sharing a layout.
///
/// The color-bearing fields (`placeholders` colors/fills/strokes/bullets +
/// `background`) are resolved against the `theme` passed to `parse_layout`. For
/// the common no-`clrMapOvr` slide that theme is the master's baked theme, so
/// the cached instance is reused; a slide with a `<p:clrMapOvr>` builds a fresh
/// `ParsedLayout` against its override theme (see the `parse_presentation` loop)
/// so its layout colors flip too. The layout's DECORATIVE spTree shapes are NOT
/// held here — they are walked per-slide because they resolve against the slide's
/// own `smartart_drawings` (§19.3.1.39 layout decorations) and are theme+zip
/// bound; caching them keyed by layout would be unsound.
#[derive(serde::Serialize)]
pub(crate) struct ParsedLayout {
    pub(crate) placeholders: LayoutPlaceholders,
    /// Layout-level `<p:cSld><p:bg>` fill (ECMA-376 §19.3.1.1 / §20.1.8.14),
    /// resolved against `theme`. Applied by the slide only when its own bg chain
    /// (slide-level) resolves to nothing.
    pub(crate) background: Option<Fill>,
    /// The LAYOUT's own `showMasterSp` (§19.3.1.39). The slide ANDs this with its
    /// own slide-level flag before compositing master decorations.
    pub(crate) show_master_sp: bool,
}

impl Default for ParsedLayout {
    fn default() -> Self {
        // Matches the prior "no/unparseable layout" behaviour: no placeholders,
        // no layout background, and showMasterSp defaulting to true.
        ParsedLayout {
            placeholders: LayoutPlaceholders::default(),
            background: None,
            show_master_sp: true,
        }
    }
}

impl ParsedLayout {
    /// No (usable) layout: no placeholder inheritance, but ordinary text still
    /// takes the presentation defaultTextStyle levels.
    pub(crate) fn without_layout(default_text: &DefaultTextLevels) -> Self {
        ParsedLayout {
            placeholders: LayoutPlaceholders {
                default_text: default_text.clone(),
                ..LayoutPlaceholders::default()
            },
            ..ParsedLayout::default()
        }
    }
}

/// ECMA-376 §19.3.1.38/§19.3.1.39 showMasterSp: absent / "1" / "true" ⇒ true;
/// "0" / "false" ⇒ false. Read from a slide or layout root element.
pub(crate) fn read_show_master_sp(node: roxmltree::Node<'_, '_>) -> bool {
    match attr(&node, "showMasterSp").as_deref() {
        Some("0") | Some("false") => false,
        _ => true, // default true (absent / "1" / "true")
    }
}

/// Parse a slide layout's XML EXACTLY ONCE and extract everything a slide
/// inherits from it (D4). Replaces the four former per-slide layout
/// `Document::parse` calls in `parse_slide` (placeholders, background,
/// showMasterSp, decorations) — the decorations still walk per-slide, but from
/// the SAME `Document` when the caller reuses it, and the other three are cached.
/// `theme` is the slide's effective theme (master-baked, or override-adjusted);
/// the master maps are the inheritance fallbacks, threaded through unchanged.
#[allow(clippy::too_many_arguments)]
pub(crate) fn parse_layout(
    layout_xml: &str,
    master_level_faces: &HashMap<String, LevelFaces>,
    default_text: &DefaultTextLevels,
    master_styles: &MasterStyleTier,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_colors: &HashMap<String, LevelColors>,
    master_level_run_properties: &MasterLevelRunProperties,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_body_pr: &HashMap<String, InheritedBodyPr>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_spacing: &HashMap<String, LevelSpacing>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    layout_part: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> ParsedLayout {
    note_layout_master_parse();
    let doc = match parse_preflighted_pptx_xml(layout_xml) {
        Ok(d) => d,
        // Unparseable layout → same as no layout: default placeholders/bg and
        // showMasterSp = true (the slide's own flag still applies downstream).
        Err(_) => return ParsedLayout::without_layout(default_text),
    };
    let root = doc.root_element();

    let placeholders = parse_layout_placeholders(
        root,
        master_level_faces,
        default_text,
        master_styles,
        master_level_font_sizes,
        master_level_colors,
        master_level_run_properties,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_body_pr,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_spacing,
        theme_source,
        layout_part,
        layout_rels,
        zip,
    );

    // Layout-level bg (rels = layout rels, part dir = layout_part). Verbatim from
    // the former inline layout-bg block in `parse_slide`; the slide decides
    // whether to use it (only when its own bg chain is empty).
    let background: Option<Fill> = child(root, "cSld").and_then(|n| {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = layout_rels.get(rid)?;
            let path = resolve_path(layout_part, target);
            // Existence check only — central-directory lookup, no inflate.
            zip.index_for_name(&path)?;
            Some(path)
        };
        parse_background(n, theme_source, &mut resolve)
    });

    let show_master_sp = read_show_master_sp(root);

    ParsedLayout {
        placeholders,
        background,
        show_master_sp,
    }
}

/// All slide-master-derived data plus the master's effective theme, bundled so
/// it can be computed once per master and reused across every slide that shares
/// that master (ECMA-376 §19.3.1.42 — a deck may have multiple masters, each
/// with its own theme/clrMap). Resolving theme/master per slide via the
/// slide→slideLayout→slideMaster→theme rels chain is required so that scheme
/// colors (e.g. `<a:schemeClr val="accent1">`) pick the right palette.
#[derive(serde::Serialize)]
pub(crate) struct ParsedMaster {
    /// The master's effective theme palette, with the master's `<p:clrMap>`
    /// pre-baked (logical names → slot hex). Includes font/line
    /// keys exactly as `parse_theme_colors` produced them.
    pub(crate) theme: PptxTheme,
    pub(crate) master_xml: Option<String>,
    pub(crate) master_rels: HashMap<String, String>,
    pub(crate) master_part: String,
    pub(crate) master_smartart_drawings: HashMap<String, String>,
    pub(crate) master_bg: Option<Fill>,
    /// The master's own decorative (non-placeholder) spTree shapes, resolved ONCE
    /// against the master's baked `theme` (§19.3.1.38 showMasterSp). Each slide
    /// composites these beneath its content; pre-extracting here (per cached
    /// master) removes the per-slide master-XML re-parse + spTree re-walk (D4).
    /// A slide with a `<p:clrMapOvr>` re-resolves them against its override theme
    /// (see `parse_slide`), so these frozen-against-master-theme elements are used
    /// only by the common no-override slides.
    pub(crate) master_decorative: Vec<SlideElement>,
    pub(crate) master_level_faces: HashMap<String, LevelFaces>,
    /// Presentation defaultTextStyle levels resolved against this master's theme.
    pub(crate) default_text: DefaultTextLevels,
    /// txStyles-only tier, resolved against this master's theme.
    pub(crate) master_styles: MasterStyleTier,
    pub(crate) master_level_font_sizes: HashMap<String, LevelFontSizes>,
    pub(crate) master_level_colors: HashMap<String, LevelColors>,
    pub(crate) master_level_run_properties: MasterLevelRunProperties,
    pub(crate) master_level_indents: HashMap<String, LevelIndents>,
    pub(crate) master_level_bullets: HashMap<String, LevelBullets>,
    pub(crate) master_anchors: HashMap<String, String>,
    pub(crate) master_body_pr: HashMap<String, InheritedBodyPr>,
    pub(crate) master_transforms: HashMap<String, Transform>,
    pub(crate) master_alignments: HashMap<String, String>,
    pub(crate) master_ea_ln_brk: HashMap<String, bool>,
    pub(crate) master_spacing: HashMap<String, LevelSpacing>,
    pub(crate) master_bold: HashMap<String, bool>,
    pub(crate) master_italic: HashMap<String, bool>,
    pub(crate) master_caps: HashMap<String, String>,
    pub(crate) master_reflection: HashMap<String, Reflection>,
    pub(crate) master_color: HashMap<String, String>,
}

/// The subset of `ParsedMaster` fields that are THEME-DEPENDENT, recomputed for a
/// slide whose `<p:clrMapOvr><a:overrideClrMapping>` (ECMA-376 §19.3.1.7) replaces
/// the master's color mapping for the WHOLE slide (§20.1.6.8). `build_master_bundle`
/// freezes these against the MASTER's own clrMap-baked theme; for an override slide
/// we re-resolve them against the slide's effective mapping so that master-INHERITED
/// scheme colors (a `<p:bg>` schemeClr, master txStyles placeholder colors, master
/// bullet colors) flip together with the slide's own shapes. Owns all its data and
/// holds no `zip` borrow, so it can be built before `parse_slide(zip)` is called.
pub(crate) struct EffectiveMaster {
    /// `bundle.theme` clone with the override mapping applied (logical → slot hex).
    pub(crate) theme: PptxTheme,
    /// Master `<p:bg>` re-resolved against `theme` (replaces `ParsedMaster.master_bg`).
    pub(crate) master_bg: Option<Fill>,
    /// Master txStyles placeholder colors re-resolved against `theme`.
    pub(crate) master_color: HashMap<String, String>,
    /// Master list-level colours re-resolved against `theme`.
    pub(crate) master_level_colors: HashMap<String, LevelColors>,
    pub(crate) master_level_run_properties: MasterLevelRunProperties,
    /// Master per-level bullet colors re-resolved against `theme`.
    pub(crate) master_level_bullets: HashMap<String, LevelBullets>,
    /// txStyles-only tier re-resolved against `theme`.
    pub(crate) master_styles: MasterStyleTier,
    /// Presentation defaultTextStyle levels with their colours re-resolved
    /// against `theme` (`default_text_level_colors`).
    pub(crate) default_text: DefaultTextLevels,
}

/// Build a `ParsedMaster` for the master at `master_path` (a ZIP path such as
/// `ppt/slideMasters/slideMaster2.xml`). Reads the master XML + its rels,
/// resolves the master's own `/theme` relationship, parses the theme colors,
/// bakes the master's `<p:clrMap>`, then computes every master-derived map.
///
/// `fallback_theme` is the presentation-level theme used only when the master
/// has no `/theme` relationship of its own (keeps simple single-theme decks and
/// malformed packages working).
///
/// A slide/layout theme override supplies `effective_theme` before extraction:
/// fonts and style-matrix recipes in inherited master content must be resolved
/// again as well as colors (CT_BaseStylesOverride, §20.1.6.12).
pub(crate) fn build_master_bundle(
    master_path: &str,
    fallback_theme: &PptxTheme,
    effective_theme: Option<&PptxTheme>,
    default_text_style_xml: Option<&str>,
    zip: &mut PptxZip,
) -> ParsedMaster {
    let master_xml_opt: Option<String> = if master_path.is_empty() {
        None
    } else {
        read_zip_str(zip, master_path).ok()
    };

    let master_part = if master_path.is_empty() {
        "ppt/slideMasters/slideMaster.xml".to_owned()
    } else {
        master_path.to_owned()
    };

    // Master rels: `<master_part>/_rels/<file>.rels`.
    let master_rels_xml: String = if master_path.is_empty() {
        // An empty path is the explicit no-master fallback, not an OPC source
        // part. It has no relationship part; deriving `_rels/.rels` would read
        // an unrelated/malicious package entry into the fallback inheritance.
        String::new()
    } else {
        let rels_p = relationship_part_path(master_path);
        read_zip_str(zip, &rels_p).unwrap_or_default()
    };
    let master_rels: HashMap<String, String> = parse_rels(&master_rels_xml);

    // The master's own theme (slide→…→slideMaster→theme). Fall back to the
    // presentation theme when the master declares no /theme relationship.
    let theme_path: Option<String> =
        find_rel_target_by_type(&master_rels_xml, "/theme").map(|t| resolve_path(&master_part, &t));
    let mut theme = effective_theme.cloned().unwrap_or_else(|| {
        theme_path
            .as_deref()
            .map(|path| parse_theme_part(path, zip))
            .unwrap_or_else(|| fallback_theme.clone())
    });
    // Bake the master's <p:clrMap> logical-name → slot mapping into the theme.
    bake_clr_map(&mut theme, master_xml_opt.as_deref());

    let master_smartart_drawings: HashMap<String, String> =
        build_smartart_drawings(&master_rels_xml, &master_part, zip);

    // Parse the master XML EXACTLY ONCE and share the resulting `Document` across
    // every master-derived extractor below (D4: previously each `parse_master_*`
    // re-ran `Document::parse` on the same string, so a single master cost 12
    // parses — 11 extractors + the background). The `Document` borrows
    // `master_xml_opt`, so it lives only for the extraction scope; all owned maps
    // are computed before it is dropped. When the master has no XML (missing part)
    // every map defaults to empty, exactly as the prior `Option::map` chain did.
    let master_doc: Option<roxmltree::Document<'_>> = master_xml_opt.as_deref().and_then(|xml| {
        note_layout_master_parse();
        parse_preflighted_pptx_xml(xml).ok()
    });
    let master_root: Option<roxmltree::Node<'_, '_>> =
        master_doc.as_ref().map(|d| d.root_element());

    let master_bg: Option<Fill> = master_root.and_then(|root| {
        let c_sld = child(root, "cSld")?;
        let mut resolve = |rid: &str| -> Option<String> {
            let target = master_rels.get(rid)?;
            let path = resolve_path(&master_part, target);
            // Existence check only — central-directory lookup, no inflate
            // (former `read_zip_bytes` decompressed the entry just to discard it).
            zip.index_for_name(&path)?;
            Some(path)
        };
        parse_background(c_sld, &theme, &mut resolve)
    });

    // The presentation `<p:defaultTextStyle>` (see `default_text_style_fragment`).
    let dts_doc: Option<roxmltree::Document<'_>> =
        default_text_style_xml.and_then(|xml| parse_preflighted_pptx_xml(xml).ok());
    let dts: Option<roxmltree::Node<'_, '_>> = dts_doc.as_ref().and_then(|doc| {
        doc.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "defaultTextStyle")
    });
    let mut default_text = parse_default_text_levels(dts, &theme);
    if let Some(root) = master_root {
        default_text.table = parse_table_text_levels(root);
    }
    let master_styles = master_root
        .map(|root| MasterStyleTier::parse(root, &theme, &master_rels, &master_part, dts, zip))
        .unwrap_or_default();
    let master_level_faces = master_root
        .map(|root| parse_master_level_faces(root, &theme, dts))
        .unwrap_or_default();
    let master_level_font_sizes = master_root
        .map(|root| parse_master_level_font_sizes(root, dts))
        .unwrap_or_default();
    let master_level_colors = master_root
        .map(|root| parse_master_level_colors(root, &theme))
        .unwrap_or_default();
    let master_level_run_properties = master_root
        .map(|root| parse_master_level_run_properties(root, &theme, &master_rels, &master_part))
        .unwrap_or_default();
    let master_level_indents = master_root
        .map(|root| parse_master_level_indents(root))
        .unwrap_or_default();
    let master_level_bullets = master_root
        .map(|root| parse_master_level_bullets(root, &theme, &master_rels, &master_part, zip))
        .unwrap_or_default();
    let master_anchors = master_root.map(parse_master_anchors).unwrap_or_default();
    let master_body_pr = master_root
        .map(parse_master_text_body_properties)
        .unwrap_or_default();
    let master_transforms = master_root.map(parse_master_transforms).unwrap_or_default();
    let master_alignments = master_root
        .map(|root| parse_master_alignments(root))
        .unwrap_or_default();
    let master_ea_ln_brk = master_root.map(parse_master_ea_ln_brk).unwrap_or_default();
    let master_spacing = master_root
        .map(|root| parse_master_txstyle_spacing(root))
        .unwrap_or_default();
    let (master_bold, master_italic, master_caps, master_reflection) = master_root
        .map(|root| parse_master_txstyle_run_properties(root))
        .unwrap_or_default();
    let master_color = master_root
        .map(|root| parse_master_txstyle_color(root, &theme))
        .unwrap_or_default();

    // Pre-extract the master's decorative (non-placeholder) spTree shapes ONCE,
    // resolved against the master's baked `theme`. Each slide clones these instead
    // of re-parsing the master XML and re-walking its spTree (D4; former
    // per-slide `parse_slide` inline walk). Uses the same shared `master_root` and
    // the master's own rels + smartart drawings, exactly as the old inline walk did.
    let mut master_decorative: Vec<SlideElement> = Vec::new();
    if let Some(root) = master_root {
        extract_decorative_shapes(
            root,
            &master_part,
            &master_rels,
            &master_smartart_drawings,
            &theme,
            &default_text,
            zip,
            &mut master_decorative,
        );
    }

    ParsedMaster {
        theme,
        master_xml: master_xml_opt,
        master_rels,
        master_part,
        master_smartart_drawings,
        master_bg,
        master_decorative,
        master_level_faces,
        default_text,
        master_styles,
        master_level_font_sizes,
        master_level_colors,
        master_level_run_properties,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_body_pr,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_spacing,
        master_bold,
        master_italic,
        master_caps,
        master_reflection,
        master_color,
    }
}

#[cfg(test)]
mod placeholder_geometry_tests {
    use super::*;
    use crate::shape::parse_shape;
    use crate::text::{BuMarker, BulletProps, ParagraphSpacing};
    use std::io::Cursor;

    fn empty_zip() -> PptxZip {
        let writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let cursor = writer.finish().unwrap();
        PptxZip::new(cursor).unwrap()
    }

    #[test]
    fn layout_partial_run_defaults_keep_master_pattern_and_other_fields() {
        let master = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld><p:txStyles>
            <p:bodyStyle><a:lvl1pPr><a:defRPr b="0" i="1" sz="3400">
              <a:pattFill prst="dnDiag"><a:fgClr><a:srgbClr val="D21D54"/></a:fgClr>
                <a:bgClr><a:srgbClr val="12CED4"/></a:bgClr></a:pattFill>
            </a:defRPr></a:lvl1pPr></p:bodyStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let layout = r#"<p:sldLayout
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree><p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/>
              <p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
            <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>
              <a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr>
            </a:lstStyle><a:p/></p:txBody>
          </p:sp></p:spTree></p:cSld>
        </p:sldLayout>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let theme = HashMap::new();
        let master_runs = parse_master_level_run_properties(
            master_doc.root_element(),
            &theme,
            &HashMap::new(),
            "ppt/slideMasters/slideMaster1.xml",
        );
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &crate::master::DefaultTextLevels::default(),
            &crate::master::MasterStyleTier::default(),
            &HashMap::new(),
            &HashMap::new(),
            &master_runs,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &theme,
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let level = placeholders.lookup_level_run_properties("body", Some(1));
        let value = serde_json::to_value(&level[0]).unwrap();
        assert_eq!(value["bold"], true);
        assert_eq!(value["italic"], true);
        assert_eq!(value["font_size"], 34.0);
        assert!(value["fill"].to_string().contains("dnDiag"));
    }

    fn parse_layout_with_master(
        layout_shape: &str,
        master_level_sizes: &HashMap<String, LevelFontSizes>,
    ) -> LayoutPlaceholders {
        let xml = format!(
            r#"<p:sldLayout
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>{layout_shape}</p:spTree></p:cSld>
                </p:sldLayout>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let mut zip = empty_zip();
        parse_layout_placeholders(
            doc.root_element(),
            &HashMap::<String, LevelFaces>::new(),
            &DefaultTextLevels::default(),
            &MasterStyleTier {
                sizes: master_level_sizes.clone(),
                ..Default::default()
            },
            master_level_sizes,
            &HashMap::<String, LevelColors>::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::<String, LevelIndents>::new(),
            &HashMap::<String, LevelBullets>::new(),
            &HashMap::<String, String>::new(),
            &HashMap::<String, InheritedBodyPr>::new(),
            &HashMap::<String, Transform>::new(),
            &HashMap::<String, String>::new(),
            &HashMap::<String, bool>::new(),
            &HashMap::<String, LevelSpacing>::new(),
            &HashMap::new(),
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        )
    }

    fn parse_layout_geometry(layout_shape: &str) -> LayoutPlaceholders {
        parse_layout_with_master(layout_shape, &HashMap::new())
    }

    fn parse_slide_shape(shape: &str, placeholders: &LayoutPlaceholders) -> ShapeElement {
        let xml = format!(
            r#"<p:sp
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  {shape}
                </p:sp>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let mut zip = empty_zip();
        parse_shape(
            doc.root_element(),
            placeholders,
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slides/slide1.xml",
            None,
            &mut zip,
        )
        .unwrap()
    }

    /// ECMA-376 makes p:txStyles optional on a slide master. PowerPoint still
    /// applies its presentation placeholder defaults when that authored tier is
    /// absent. The values below are bounded to an Office-produced matrix that
    /// distinguishes title, body/subtitle, and object placeholders. These are
    /// application defaults, not fabricated defaults for ordinary text boxes.
    /// ECMA-376 §19.3.1.51 txStyles: a placeholder without its own or a layout
    /// lnSpc inherits the master lnSpc of its type, level by level. An
    /// idx-matched layout placeholder without lnSpc also falls through to the
    /// master, and a level the layout does not set keeps the master's value.
    #[test]
    fn master_tx_styles_line_spacing_is_inherited() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="85000"/></a:lnSpc></a:lvl1pPr></p:titleStyle>
            <p:bodyStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:lvl1pPr>
              <a:lvl2pPr><a:lnSpc><a:spcPct val="80000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef></a:lvl2pPr></p:bodyStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let pct = |val: f64| Some(ooxml_common::text::SpaceLine::Pct { val });
        let doc = roxmltree::Document::parse(xml).unwrap();
        let spacing = parse_master_txstyle_spacing(doc.root_element());
        assert_eq!(spacing["title"].line[0], pct(85000.0));
        // The title style sets level 1 only; level 2 is not borrowed from it.
        assert_eq!(spacing["title"].line[1], None);
        assert_eq!(spacing["body"].line[..2], [pct(90000.0), pct(80000.0)]);
        assert_eq!(
            spacing["body"].before[1],
            Some(ParagraphSpacing::Points(500))
        );
        assert_eq!(spacing["obj"].line[0], pct(90000.0));
        assert!(!spacing.contains_key("dt"));

        // idx 11 and 12 are bound layout slots; an unmatched idx reads the
        // txStyles-only tier.
        let mut layout_12 = LevelSpacing::default();
        layout_12.line[0] = pct(120000.0);
        let placeholders = LayoutPlaceholders {
            by_idx_placeholder_type: HashMap::from([
                (11, "body".to_owned()),
                (12, "body".to_owned()),
            ]),
            by_idx_spacing: HashMap::from([(12, layout_12)]),
            by_type_master_spacing: spacing.clone(),
            styles: MasterStyleTier {
                spacing,
                ..Default::default()
            },
            ..LayoutPlaceholders::default()
        };
        assert_eq!(
            placeholders.lookup_spacing("body", Some(40)).line[0],
            pct(90000.0)
        );
        assert_eq!(
            placeholders.lookup_spacing("body", Some(11)).line[0],
            pct(90000.0)
        );
        let slot_12 = placeholders.lookup_spacing("body", Some(12));
        assert_eq!(slot_12.line[..2], [pct(120000.0), pct(80000.0)]);
        assert_eq!(
            placeholders.lookup_spacing("title", None).line[0],
            pct(85000.0)
        );
        assert!(placeholders.lookup_spacing("dt", Some(3)).is_empty());
    }

    /// Review regression (#1630): list-level indents and spacing merge per
    /// level and per axis. A master placeholder that indents only level 1 keeps
    /// the built-in body style's other levels, and an authored point value
    /// (`spcPts`) of lnSpc / spcBef / spcAft survives the cascade.
    #[test]
    fn partial_master_list_levels_merge_per_level_and_axis() {
        use ooxml_common::text::SpaceLine;
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree>
            <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
              <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>
                <a:lvl1pPr marL="342900"><a:spcAft><a:spcPct val="20000"/></a:spcAft></a:lvl1pPr>
                <a:lvl3pPr><a:lnSpc><a:spcPts val="3000"/></a:lnSpc><a:spcBef><a:spcPts val="700"/></a:spcBef></a:lvl3pPr>
              </a:lstStyle><a:p/></p:txBody></p:sp>
          </p:spTree></p:cSld>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let indents = parse_master_level_indents(doc.root_element());
        let body = indents["body"];
        // Level 1: marL from the placeholder, indent from the built-in style.
        assert_eq!(
            (body[0].mar_l, body[0].indent),
            (Some(342_900), Some(-228_600))
        );
        // Level 3 untouched by the placeholder: the built-in (90 pt, -18 pt).
        assert_eq!(
            (body[2].mar_l, body[2].indent),
            (Some(1_143_000), Some(-228_600))
        );

        let spacing = parse_master_txstyle_spacing(doc.root_element());
        let body = &spacing["body"];
        assert_eq!(body.line[2], Some(SpaceLine::Pts { val: 30.0 }));
        assert_eq!(body.before[2], Some(ParagraphSpacing::Points(700)));
        assert_eq!(body.after[0], Some(ParagraphSpacing::Percent(20000.0)));
        // Unset properties keep the built-in values per level.
        assert_eq!(body.line[0], Some(SpaceLine::Pct { val: 90000.0 }));
        assert_eq!(body.before[0], Some(ParagraphSpacing::Points(1000)));
        assert_eq!(body.line[1], Some(SpaceLine::Pct { val: 90000.0 }));
    }

    const PML_A: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    fn parse_shape_with_theme(
        shape: &str,
        placeholders: &LayoutPlaceholders,
        theme: &HashMap<String, String>,
    ) -> ShapeElement {
        let xml = format!(r#"<p:sp {PML_A}>{shape}</p:sp>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let mut zip = empty_zip();
        parse_shape(
            doc.root_element(),
            placeholders,
            theme,
            &HashMap::new(),
            "ppt/slides/slide1.xml",
            None,
            &mut zip,
        )
        .unwrap()
    }

    fn paragraph_faces(shape: &ShapeElement) -> Vec<(Option<String>, Option<f64>, bool)> {
        shape
            .text_body
            .as_ref()
            .unwrap()
            .paragraphs
            .iter()
            .map(|p| {
                (
                    p.def_font_family.clone(),
                    p.def_font_size,
                    matches!(p.bullet, Bullet::Char { .. }),
                )
            })
            .collect()
    }

    /// Issue #1620 (PowerPoint for Mac PDF export): a slide placeholder bound
    /// to a layout slot WITHOUT a txBody renders in Arial 18 pt with no bullet
    /// even though the master has txStyles; a slot with a txBody and an idx with
    /// no slot both reach the master body style.
    #[test]
    fn layout_slot_without_text_body_cuts_master_list_styles() {
        let master = format!(
            r#"<p:sldMaster {PML_A}><p:cSld><p:spTree/></p:cSld><p:txStyles>
              <p:titleStyle><a:lvl1pPr><a:defRPr sz="4000"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>
              <p:bodyStyle>
                <a:lvl1pPr><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="2400"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>
                <a:lvl2pPr><a:buChar char="&#8226;"/><a:defRPr sz="2000"><a:latin typeface="Courier New"/></a:defRPr></a:lvl2pPr>
              </p:bodyStyle></p:txStyles></p:sldMaster>"#
        );
        let theme = HashMap::from([
            ("+mj-lt".to_owned(), "Georgia".to_owned()),
            ("+mn-lt".to_owned(), "Verdana".to_owned()),
        ]);
        let master_doc = roxmltree::Document::parse(&master).unwrap();
        let root = master_doc.root_element();
        let layout = format!(
            r#"<p:sldLayout {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="B1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="4" name="B2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
            </p:spTree></p:cSld></p:sldLayout>"#
        );
        let layout_doc = roxmltree::Document::parse(&layout).unwrap();
        let mut zip = empty_zip();
        let styles = MasterStyleTier::parse(
            root,
            &theme,
            &HashMap::new(),
            "ppt/slideMasters/slideMaster1.xml",
            None,
            &mut zip,
        );
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &parse_master_level_faces(root, &theme, None),
            &DefaultTextLevels::default(),
            &styles,
            &parse_master_level_font_sizes(root, None),
            &HashMap::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &parse_master_level_bullets(
                root,
                &theme,
                &HashMap::new(),
                "ppt/slideMasters/slideMaster1.xml",
                &mut zip,
            ),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &theme,
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let body = |ph: &str| {
            format!(
                r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/>
                  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>a</a:t></a:r></a:p>
                  <a:p><a:pPr lvl="1"/><a:r><a:t>b</a:t></a:r></a:p></p:txBody>"#
            )
        };
        let arial = |size: f64| (Some("Arial".to_owned()), Some(size), false);

        let cut = parse_shape_with_theme(
            &body(r#"<p:ph type="body" idx="1"/>"#),
            &placeholders,
            &theme,
        );
        assert_eq!(paragraph_faces(&cut), vec![arial(18.0), arial(18.0)]);
        let title = parse_shape_with_theme(&body(r#"<p:ph type="title"/>"#), &placeholders, &theme);
        assert_eq!(paragraph_faces(&title), vec![arial(18.0), arial(18.0)]);

        let styled = vec![
            (Some("Verdana".to_owned()), Some(24.0), true),
            (Some("Courier New".to_owned()), Some(20.0), true),
        ];
        let with_body = parse_shape_with_theme(
            &body(r#"<p:ph type="body" idx="2"/>"#),
            &placeholders,
            &theme,
        );
        assert_eq!(paragraph_faces(&with_body), styled);
        let no_slot = parse_shape_with_theme(
            &body(r#"<p:ph type="body" idx="7"/>"#),
            &placeholders,
            &theme,
        );
        assert_eq!(paragraph_faces(&no_slot), styled);
    }

    /// Issue #1620: levels do not inherit from level 1; a present level with no
    /// latin and an absent level both end at Arial (18 pt when no size), and a
    /// theme token naming an empty slot falls through like an absent face.
    #[test]
    fn placeholder_levels_end_at_the_hard_default() {
        let master = format!(
            r#"<p:sldMaster {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Constantia"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld><p:txStyles>
              <p:titleStyle><a:lvl1pPr><a:defRPr sz="4000"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>
              <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>
                <a:lvl3pPr><a:defRPr sz="1600"/></a:lvl3pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#
        );
        // An empty major font: `+mj-lt` is unspecified, not "".
        let theme = HashMap::from([
            ("+mj-lt".to_owned(), String::new()),
            ("+mn-lt".to_owned(), "Candara".to_owned()),
        ]);
        let doc = roxmltree::Document::parse(&master).unwrap();
        let faces = parse_master_level_faces(doc.root_element(), &theme, None);
        let sizes = parse_master_level_font_sizes(doc.root_element(), None);
        let placeholders = LayoutPlaceholders {
            by_type_level_faces: faces.clone(),
            by_type_master_level_faces: faces,
            by_type_level_sizes: sizes.clone(),
            by_type_master_level_sizes: sizes,
            ..Default::default()
        };
        let text = |ph: &str| {
            format!(
                r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/>
                  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>1</a:t></a:r></a:p>
                  <a:p><a:pPr lvl="1"/><a:r><a:t>2</a:t></a:r></a:p>
                  <a:p><a:pPr lvl="2"/><a:r><a:t>3</a:t></a:r></a:p></p:txBody>"#
            )
        };
        let faces_sizes = |shape: &ShapeElement| {
            paragraph_faces(shape)
                .into_iter()
                .map(|(f, s, _)| (f.unwrap(), s.unwrap()))
                .collect::<Vec<_>>()
        };
        // The obj slot takes the master BODY placeholder's lstStyle face.
        let obj = parse_shape_with_theme(&text(r#"<p:ph idx="5"/>"#), &placeholders, &theme);
        assert_eq!(
            faces_sizes(&obj),
            vec![
                ("Constantia".to_owned(), 24.0),
                ("Arial".to_owned(), 18.0),
                ("Arial".to_owned(), 16.0),
            ]
        );
        let title = parse_shape_with_theme(&text(r#"<p:ph type="title"/>"#), &placeholders, &theme);
        assert_eq!(faces_sizes(&title)[0], ("Arial".to_owned(), 40.0));
    }

    /// Issue #1620: ordinary text takes the defaultTextStyle level, a style
    /// fontRef major/minor overrides only its face, and an explicit empty
    /// typeface is Arial rather than an inherited face.
    #[test]
    fn ordinary_text_uses_default_text_style_and_font_ref() {
        let theme = HashMap::from([
            ("+mj-lt".to_owned(), "Cambria".to_owned()),
            ("+mn-lt".to_owned(), String::new()),
        ]);
        let dts = format!(
            r#"<p:defaultTextStyle {PML_A}>
              <a:lvl1pPr><a:defRPr sz="2000"><a:latin typeface="Century Gothic"/></a:defRPr></a:lvl1pPr>
              <a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr></p:defaultTextStyle>"#
        );
        let dts_doc = roxmltree::Document::parse(&dts).unwrap();
        let placeholders = LayoutPlaceholders {
            default_text: parse_default_text_levels(Some(dts_doc.root_element()), &theme),
            ..Default::default()
        };
        let shape = |style: &str, runs: &str| {
            format!(
                r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>{style}
                  <p:txBody><a:bodyPr/><a:lstStyle/>{runs}</p:txBody>"#
            )
        };
        let two_levels = r#"<a:p><a:r><a:t>1</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:t>2</a:t></a:r></a:p>"#;
        let text_box = parse_shape_with_theme(&shape("", two_levels), &placeholders, &theme);
        assert_eq!(
            paragraph_faces(&text_box),
            vec![
                (Some("Century Gothic".to_owned()), Some(20.0), false),
                (Some("Arial".to_owned()), Some(20.0), false),
            ]
        );
        let font_ref = |idx: &str| {
            format!(
                r#"<p:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="{idx}"/></p:style>"#
            )
        };
        let major = parse_shape_with_theme(
            &shape(&font_ref("major"), two_levels),
            &placeholders,
            &theme,
        );
        assert_eq!(
            paragraph_faces(&major)
                .into_iter()
                .map(|(f, s, _)| (f.unwrap(), s.unwrap()))
                .collect::<Vec<_>>(),
            vec![("Cambria".to_owned(), 20.0), ("Cambria".to_owned(), 20.0)]
        );
        let none =
            parse_shape_with_theme(&shape(&font_ref("none"), two_levels), &placeholders, &theme);
        assert_eq!(paragraph_faces(&none), paragraph_faces(&text_box));

        let runs = r#"<a:p><a:r><a:rPr><a:latin typeface=""/></a:rPr><a:t>e</a:t></a:r>
          <a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>t</a:t></a:r></a:p>"#;
        let body = parse_shape_with_theme(&shape("", runs), &placeholders, &theme)
            .text_body
            .unwrap();
        let paragraph = &body.paragraphs[0];
        let run_faces: Vec<_> = paragraph
            .runs
            .iter()
            .map(|run| match run {
                TextRun::Text(t) => t.font_family.clone(),
                _ => None,
            })
            .collect();
        // "" is Arial; `+mn-lt` naming an empty slot is unspecified and falls
        // to the paragraph's defaultTextStyle face.
        assert_eq!(
            run_faces,
            vec![Some("Arial".to_owned()), Some("Century Gothic".to_owned())]
        );
        assert_eq!(paragraph.def_font_family.as_deref(), Some("Century Gothic"));
    }

    #[test]
    fn master_without_tx_styles_uses_built_in_title_and_body_styles() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let theme = HashMap::from([
            ("+mj-lt".to_owned(), "Calibri Light".to_owned()),
            ("+mn-lt".to_owned(), "Calibri".to_owned()),
        ]);

        let sizes = parse_master_level_font_sizes(doc.root_element(), None);
        let faces = parse_master_level_faces(doc.root_element(), &theme, None);
        let indents = parse_master_level_indents(doc.root_element());
        let mut zip = empty_zip();
        let bullets = parse_master_level_bullets(
            doc.root_element(),
            &theme,
            &HashMap::new(),
            "ppt/slideMasters/slideMaster1.xml",
            &mut zip,
        );

        for title in TITLE_CLASS {
            assert_eq!(sizes[*title][0], Some(44.0));
            // Chain values keep theme tokens (issue #1627).
            assert_eq!(faces[*title][0].as_deref(), Some("+mj-lt"));
        }
        // Every body-class type, obj included, gets the built-in body style.
        for body in BODY_CLASS {
            assert_eq!(sizes[*body][0], Some(28.0));
            assert_eq!(faces[*body][0].as_deref(), Some("+mn-lt"));
            assert_eq!(indents[*body][0].mar_l, Some(228_600));
            assert_eq!(indents[*body][0].indent, Some(-228_600));
            match bullets[*body][0].resolve() {
                Bullet::Char {
                    ch, font_family, ..
                } => {
                    assert_eq!(ch, "•");
                    assert_eq!(font_family.as_deref(), Some("Arial"));
                }
                other => panic!("expected the built-in body bullet, got {other:?}"),
            }
        }
        // Body levels 2-9 (#1630 controls, levels 1-5 and 6-9): 24 / 20 / 18 pt
        // and 18 pt below, marL + 0.5" per level, the same hanging Arial
        // bullet, theme minor face.
        for (level, size) in [
            (1, 24.0),
            (2, 20.0),
            (3, 18.0),
            (4, 18.0),
            (5, 18.0),
            (8, 18.0),
        ] {
            assert_eq!(sizes["body"][level], Some(size));
            // The chain keeps the theme token (issue #1627).
            assert_eq!(faces["body"][level].as_deref(), Some("+mn-lt"));
            assert_eq!(
                indents["body"][level].mar_l,
                Some(228_600 + 457_200 * level as i64)
            );
            assert_eq!(indents["body"][level].indent, Some(-228_600));
            assert!(matches!(
                bullets["body"][level].resolve(),
                Bullet::Char { .. }
            ));
        }
        // The title style has level 1 only; deeper title levels end at the
        // hard defaults (Arial 18 pt, no indent).
        assert_eq!(sizes["title"][1], None);
        assert_eq!(faces["title"][1], None);
        assert!(indents
            .get("title")
            .is_none_or(|levels| levels[1].mar_l.is_none()));
        assert!(!sizes.contains_key("dt"));
    }

    /// #1435 control: on a master without txStyles, title, subTitle and body
    /// text with no fill renders the theme tx1 through the slide's colour map
    /// (the default template's fill), not the 000000 hard default.
    #[test]
    fn master_without_tx_styles_colours_title_and_body_text_tx1() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        // A dark map: tx1 resolves to lt1.
        let theme = HashMap::from([
            ("dk1".to_owned(), "10263F".to_owned()),
            ("lt1".to_owned(), "F9F9F9".to_owned()),
            ("tx1".to_owned(), "F9F9F9".to_owned()),
        ]);
        let colors = parse_master_level_colors(doc.root_element(), &theme);
        for class in TITLE_CLASS {
            assert_eq!(colors[*class][0].as_deref(), Some("F9F9F9"), "{class}");
            // Deeper title levels keep the hard default.
            assert_eq!(colors[*class][1], None, "{class}");
        }
        for class in BODY_CLASS {
            for level in 0..9 {
                assert_eq!(colors[*class][level].as_deref(), Some("F9F9F9"), "{class} {level}");
            }
        }
        // The level-1 fill stays on its level: it is no shape-wide default,
        // which would colour the deeper title levels too.
        assert!(parse_master_txstyle_color(doc.root_element(), &theme).is_empty());
    }

    /// Review regression (#1620 x #1625): only the Latin face and size of
    /// dt/ftr/sldNum switch to defaultTextStyle. Every other character
    /// property of a footer still inherits master otherStyle.
    /// #1620/#1992 controls: a list style's defPPr has no effect at any tier,
    /// neither its face and size nor its other character properties (bold).
    #[test]
    fn def_ppr_character_properties_are_ignored() {
        let dts = format!(
            r#"<p:defaultTextStyle {PML_A}><a:defPPr><a:defRPr sz="1900"><a:latin typeface="Perpetua"/></a:defRPr></a:defPPr>
              <a:lvl1pPr><a:defRPr sz="2000"><a:latin typeface="Century Gothic"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>"#
        );
        let dts_doc = roxmltree::Document::parse(&dts).unwrap();
        let theme = HashMap::new();
        let placeholders = LayoutPlaceholders {
            default_text: parse_default_text_levels(Some(dts_doc.root_element()), &theme),
            ..Default::default()
        };
        let text_box = parse_shape_with_theme(
            r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
              <p:txBody><a:bodyPr/><a:lstStyle><a:defPPr><a:defRPr sz="1500" b="1"><a:latin typeface="Garamond"/></a:defRPr></a:defPPr>
                <a:lvl2pPr><a:defRPr><a:latin typeface="Corbel"/></a:defRPr></a:lvl2pPr></a:lstStyle>
                <a:p><a:r><a:t>1</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:t>2</a:t></a:r></a:p></p:txBody>"#,
            &placeholders,
            &theme,
        );
        let runs: Vec<_> = text_box
            .text_body
            .unwrap()
            .paragraphs
            .iter()
            .map(|p| match &p.runs[0] {
                TextRun::Text(t) => (t.font_family.clone(), t.font_size, t.bold),
                _ => panic!("text run expected"),
            })
            .collect();
        assert_eq!(
            runs,
            vec![
                (Some("Century Gothic".to_owned()), Some(20.0), None),
                (Some("Corbel".to_owned()), Some(18.0), None),
            ]
        );
    }

    /// #1620 controls: obj and typeless slots bound to a layout slot take the
    /// master BODY placeholder's character properties; an idx with no layout
    /// slot (body included) takes only txStyles.
    #[test]
    fn obj_slots_and_unmatched_idx_character_properties() {
        let master = format!(
            r#"<p:sldMaster {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr b="1" cap="all"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr b="0"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#
        );
        let layout = format!(
            r#"<p:sldLayout {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="O"/><p:cNvSpPr/><p:nvPr><p:ph type="obj" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="4" name="T"/><p:cNvSpPr/><p:nvPr><p:ph idx="3"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
            </p:spTree></p:cSld></p:sldLayout>"#
        );
        let master_doc = roxmltree::Document::parse(&master).unwrap();
        let layout_doc = roxmltree::Document::parse(&layout).unwrap();
        let theme = HashMap::new();
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &DefaultTextLevels::default(),
            &crate::master::MasterStyleTier::default(),
            &HashMap::new(),
            &HashMap::new(),
            &parse_master_level_run_properties(
                master_doc.root_element(),
                &theme,
                &HashMap::new(),
                "ppt/slideMasters/slideMaster1.xml",
            ),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &theme,
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let run = |ph: &str| {
            let shape = parse_shape_with_theme(
                &format!(
                    r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>t</a:t></a:r></a:p></p:txBody>"#
                ),
                &placeholders,
                &theme,
            );
            match &shape.text_body.unwrap().paragraphs[0].runs[0] {
                TextRun::Text(t) => (t.color.clone(), t.bold, t.caps.clone()),
                _ => panic!("text run expected"),
            }
        };
        let master_ph = (
            Some("C00000".to_owned()),
            Some(true),
            Some("all".to_owned()),
        );
        let body_style = (Some("0070C0".to_owned()), Some(false), None);
        assert_eq!(run(r#"<p:ph type="body" idx="1"/>"#), master_ph);
        assert_eq!(run(r#"<p:ph type="obj" idx="2"/>"#), master_ph);
        assert_eq!(run(r#"<p:ph idx="3"/>"#), master_ph);
        assert_eq!(run(r#"<p:ph type="body" idx="7"/>"#), body_style);
        assert_eq!(run(r#"<p:ph type="obj" idx="8"/>"#), body_style);
        assert_eq!(run(r#"<p:ph idx="9"/>"#), body_style);
    }

    /// Review regression (#1620 round 3): an idx with no layout slot reads
    /// the txStyles tier only for every per-level map, colour included; a
    /// bound slot still reads the master placeholder. Also: a footer's
    /// eaLnBrk falls back to otherStyle.
    #[test]
    fn unmatched_idx_colour_and_footer_ea_ln_brk_follow_the_style_tier() {
        let master = format!(
            r#"<p:sldMaster {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle>
              <p:otherStyle><a:lvl1pPr eaLnBrk="0"/></p:otherStyle></p:txStyles></p:sldMaster>"#
        );
        let layout = format!(
            r#"<p:sldLayout {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="B"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="F"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
            </p:spTree></p:cSld></p:sldLayout>"#
        );
        let master_doc = roxmltree::Document::parse(&master).unwrap();
        let layout_doc = roxmltree::Document::parse(&layout).unwrap();
        let root = master_doc.root_element();
        let theme = HashMap::new();
        let mut zip = empty_zip();
        let styles = MasterStyleTier::parse(
            root,
            &theme,
            &HashMap::new(),
            "ppt/slideMasters/slideMaster1.xml",
            None,
            &mut zip,
        );
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &DefaultTextLevels::default(),
            &styles,
            &HashMap::new(),
            &parse_master_level_colors(root, &theme),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &parse_master_ea_ln_brk(root),
            &HashMap::new(),
            &theme,
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let first_paragraph = |ph: &str| {
            let shape = parse_shape_with_theme(
                &format!(
                    r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>t</a:t></a:r></a:p></p:txBody>"#
                ),
                &placeholders,
                &theme,
            );
            shape.text_body.unwrap().paragraphs.remove(0)
        };
        assert_eq!(
            first_paragraph(r#"<p:ph type="body" idx="1"/>"#)
                .def_color
                .as_deref(),
            Some("C00000")
        );
        assert_eq!(
            first_paragraph(r#"<p:ph type="body" idx="7"/>"#).def_color,
            None
        );
        assert!(!first_paragraph(r#"<p:ph type="ftr" idx="11"/>"#).ea_ln_brk);
    }

    #[test]
    fn footer_keeps_other_style_character_properties() {
        let master = format!(
            r#"<p:sldMaster {PML_A}><p:cSld><p:spTree/></p:cSld><p:txStyles>
              <p:otherStyle><a:lvl1pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
                <a:latin typeface="Trebuchet MS"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>"#
        );
        let dts = format!(
            r#"<p:defaultTextStyle {PML_A}><a:lvl1pPr><a:defRPr sz="2000"><a:latin typeface="Century Gothic"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>"#
        );
        let master_doc = roxmltree::Document::parse(&master).unwrap();
        let dts_doc = roxmltree::Document::parse(&dts).unwrap();
        let root = master_doc.root_element();
        let dts = Some(dts_doc.root_element());
        let theme = HashMap::new();
        let layout = format!(
            r#"<p:sldLayout {PML_A}><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="F"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
            </p:spTree></p:cSld></p:sldLayout>"#
        );
        let layout_doc = roxmltree::Document::parse(&layout).unwrap();
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &parse_master_level_faces(root, &theme, dts),
            &DefaultTextLevels::default(),
            &crate::master::MasterStyleTier::default(),
            &parse_master_level_font_sizes(root, dts),
            &HashMap::new(),
            &parse_master_level_run_properties(
                root,
                &theme,
                &HashMap::new(),
                "ppt/slideMasters/slideMaster1.xml",
            ),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &theme,
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let footer = parse_shape_with_theme(
            r#"<p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>f</a:t></a:r></a:p></p:txBody>"#,
            &placeholders,
            &theme,
        );
        let paragraph = &footer.text_body.unwrap().paragraphs[0];
        let TextRun::Text(run) = &paragraph.runs[0] else {
            panic!("expected a text run");
        };
        assert_eq!(run.bold, Some(true));
        assert_eq!(run.color.as_deref(), Some("C00000"));
        assert_eq!(run.font_family.as_deref(), Some("Century Gothic"));
        assert_eq!(run.font_size, Some(20.0));
    }

    #[test]
    fn other_placeholders_take_default_text_style_not_other_style() {
        let master = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:titleStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></p:titleStyle>
            <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle>
            <p:otherStyle><a:lvl1pPr><a:defRPr sz="1600"><a:latin typeface="Trebuchet MS"/></a:defRPr></a:lvl1pPr></p:otherStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let dts = r#"<p:defaultTextStyle
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:lvl1pPr><a:defRPr sz="2000"><a:latin typeface="Century Gothic"/></a:defRPr></a:lvl1pPr>
          <a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr>
        </p:defaultTextStyle>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let dts_doc = roxmltree::Document::parse(dts).unwrap();
        let root = master_doc.root_element();
        let dts = Some(dts_doc.root_element());
        let theme = HashMap::new();

        let sizes = parse_master_level_font_sizes(root, dts);
        let faces = parse_master_level_faces(root, &theme, dts);
        assert_eq!(sizes["title"][0], Some(36.0));
        assert_eq!(sizes["body"][0], Some(24.0));
        for other in OTHER_CLASS {
            assert_eq!(sizes[*other][0], Some(20.0));
            assert_eq!(faces[*other][0].as_deref(), Some("Century Gothic"));
            // A present level without <a:latin> has no face (hard default).
            assert_eq!(sizes[*other][1], Some(20.0));
            assert_eq!(faces[*other][1], None);
        }
        // Without a defaultTextStyle, otherStyle still supplies nothing.
        assert!(!parse_master_level_faces(root, &theme, None).contains_key("dt"));
        assert!(!parse_master_level_font_sizes(root, None).contains_key("dt"));
    }

    #[test]
    fn default_text_levels_synthesize_theme_minor_when_absent() {
        let theme = HashMap::from([("+mn-lt".to_owned(), "Verdana".to_owned())]);
        let absent = parse_default_text_levels(None, &theme);
        assert!(absent.faces.iter().all(|f| f.as_deref() == Some("+mn-lt")));
        assert!(absent.sizes.iter().all(|s| *s == Some(18.0)));

        let dts = r#"<p:defaultTextStyle
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr>
        </p:defaultTextStyle>"#;
        let doc = roxmltree::Document::parse(dts).unwrap();
        let present = parse_default_text_levels(Some(doc.root_element()), &theme);
        assert_eq!(
            present.faces[1], None,
            "a present level without latin is not the theme minor"
        );
        assert_eq!(present.sizes[1], Some(20.0));

        // #1620 (no defaultTextStyle): a level-2 text box paragraph started
        // 0.5" in. #1628: a present level that sets no marL starts at 0.
        assert_eq!(absent.mar_l, DEFAULT_TEXT_STYLE_MAR_L);
        assert_eq!(present.mar_l, [0; 9]);
    }

    /// #1630 / #1628: a plain paragraph whose cascade sets no marL starts at
    /// the inset in a placeholder (title levels 2-5) and at its
    /// defaultTextStyle level's marL in ordinary text.
    #[test]
    fn plain_paragraph_implicit_indent_follows_its_default_list_style() {
        let paragraph = |shape: &str, placeholders: &LayoutPlaceholders| {
            let shape = parse_slide_shape(shape, placeholders);
            shape.text_body.unwrap().paragraphs[0].mar_l
        };
        let title = r#"<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
            <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="1"/><a:r><a:t>x</a:t></a:r></a:p></p:txBody>"#;
        let text_box = r#"<p:nvSpPr><p:cNvPr id="3" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="2"/><a:r><a:t>x</a:t></a:r></a:p></p:txBody>"#;
        let mut placeholders = LayoutPlaceholders::default();
        placeholders.default_text.mar_l = DEFAULT_TEXT_STYLE_MAR_L;
        assert_eq!(paragraph(title, &placeholders), 0);
        assert_eq!(paragraph(text_box, &placeholders), 914_400);
        placeholders.default_text.mar_l = [0; 9];
        assert_eq!(paragraph(text_box, &placeholders), 0);
    }

    #[test]
    fn omitted_placeholder_type_uses_schema_default_obj() {
        let shape = parse_slide_shape(
            r#"<p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/>
                 <p:nvPr><p:ph idx="3"/></p:nvPr></p:nvSpPr>
               <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
               <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>content</a:t></a:r></a:p></p:txBody>"#,
            &LayoutPlaceholders::default(),
        );

        assert_eq!(shape.placeholder_type.as_deref(), Some("obj"));
    }

    #[test]
    fn omitted_placeholder_type_inherits_matching_layout_slot_type() {
        let placeholders = LayoutPlaceholders {
            by_idx_placeholder_type: HashMap::from([(3, "body".to_owned())]),
            ..LayoutPlaceholders::default()
        };
        let shape = parse_slide_shape(
            r#"<p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/>
                 <p:nvPr><p:ph idx="3"/></p:nvPr></p:nvSpPr>
               <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
               <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>content</a:t></a:r></a:p></p:txBody>"#,
            &placeholders,
        );

        assert_eq!(shape.placeholder_type.as_deref(), Some("body"));
    }

    #[test]
    fn object_slot_without_layout_transform_uses_master_body_box() {
        let body = Transform {
            x: 838_200,
            y: 1_825_625,
            cx: 10_515_600,
            cy: 4_351_338,
            ..Default::default()
        };
        let placeholders = LayoutPlaceholders {
            master_by_type: HashMap::from([("body".to_owned(), body.clone())]),
            ..Default::default()
        };

        let inherited = placeholders.lookup("obj", Some(1));

        assert_eq!(inherited.map(|transform| transform.x), Some(body.x));
        assert_eq!(inherited.map(|transform| transform.y), Some(body.y));
        assert_eq!(inherited.map(|transform| transform.cx), Some(body.cx));
        assert_eq!(inherited.map(|transform| transform.cy), Some(body.cy));
    }

    #[test]
    fn typeless_layout_slot_uses_schema_default_text_style() {
        let level = |size: f64| -> LevelFontSizes {
            let mut sizes = [None; 9];
            sizes[0] = Some(size);
            sizes
        };
        let master_sizes = HashMap::from([
            ("body".to_owned(), level(28.0)),
            ("obj".to_owned(), level(24.0)),
        ]);
        let placeholders = parse_layout_with_master(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/>
                 <p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
               <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#,
            &master_sizes,
        );
        let shape = parse_slide_shape(
            r#"<p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr/>
                 <p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
               <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>content</a:t></a:r></a:p></p:txBody>"#,
            &placeholders,
        );

        // CT_Placeholder defaults an omitted @type to obj. The idx is used to
        // match corresponding placeholders, but cannot rewrite that schema
        // value from an unrelated master slot that happens to reuse the same
        // idx in a two-content layout.
        assert_eq!(shape.placeholder_type.as_deref(), Some("obj"));
        assert_eq!(shape.text_body.unwrap().default_font_size, Some(24.0));
    }

    #[test]
    fn layout_retains_picture_effect_components_for_placeholder_inheritance() {
        let placeholders = parse_layout_geometry(
            r#"<p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Picture slot"/><p:cNvSpPr/>
                <p:nvPr><p:ph type="pic" idx="9"/></p:nvPr></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>
                <a:ln w="33333"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
                <a:effectLst><a:outerShdw blurRad="500" dist="700" dir="0"><a:srgbClr val="112233"/></a:outerShdw></a:effectLst>
                <a:scene3d><a:camera prst="perspectiveFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>
                <a:sp3d prstMaterial="plastic"/>
              </p:spPr>
            </p:sp>"#,
        );

        let properties = placeholders
            .lookup_picture_properties("pic", Some(9))
            .expect("layout picture properties");
        assert_eq!(
            properties.stroke.as_ref().map(|stroke| stroke.width),
            Some(33_333)
        );
        assert_eq!(
            properties.shadow.as_ref().map(|shadow| shadow.dist),
            Some(700)
        );
        assert_eq!(
            properties
                .scene3d
                .as_ref()
                .map(|scene| scene.camera.prst.as_str()),
            Some("perspectiveFront")
        );
        assert_eq!(
            properties
                .sp3d
                .as_ref()
                .map(|surface| surface.prst_material.as_str()),
            Some("plastic")
        );
    }

    #[test]
    fn blank_layout_keeps_master_body_bullet_fallback() {
        let xml = r#"<p:sldLayout
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sldLayout>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let mut zip = empty_zip();
        let mut master_bullets = HashMap::new();
        let mut body_levels = empty_level_bullets();
        body_levels[0] = BulletProps {
            marker: Some(BuMarker::Char("•".into())),
            ..Default::default()
        };
        master_bullets.insert("body".to_owned(), body_levels);

        let placeholders = parse_layout_placeholders(
            doc.root_element(),
            &HashMap::new(),
            &crate::master::DefaultTextLevels::default(),
            &crate::master::MasterStyleTier::default(),
            &HashMap::new(),
            &HashMap::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &master_bullets,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );

        match placeholders.lookup_level_bullets("body", None)[0].resolve() {
            Bullet::Char { ch, .. } => assert_eq!(ch, "•"),
            other => panic!("expected master body bullet, got {other:?}"),
        }
    }

    #[test]
    fn master_title_style_carries_run_reflection_to_title_placeholders() {
        let xml = r#"
          <p:sldMaster
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:txStyles>
              <p:titleStyle>
                <a:lvl1pPr>
                  <a:defRPr cap="all">
                    <a:effectLst>
                      <a:reflection blurRad="12700" stA="48000" endA="300"
                        endPos="55000" dir="5400000" sy="-90000"
                        algn="bl" rotWithShape="0"/>
                    </a:effectLst>
                  </a:defRPr>
                </a:lvl1pPr>
              </p:titleStyle>
            </p:txStyles>
          </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let (_, _, caps, reflections) = parse_master_txstyle_run_properties(doc.root_element());

        assert_eq!(caps.get("title").map(String::as_str), Some("all"));
        assert_eq!(caps.get("ctrTitle").map(String::as_str), Some("all"));
        for ph_type in ["title", "ctrTitle"] {
            let reflection = reflections
                .get(ph_type)
                .unwrap_or_else(|| panic!("missing reflection for {ph_type}"));
            assert_eq!(reflection.blur, 12_700);
            assert!((reflection.st_a - 0.48).abs() < 1e-9);
            assert!((reflection.end_a - 0.003).abs() < 1e-9);
            assert!((reflection.end_pos - 0.55).abs() < 1e-9);
            assert!((reflection.sy + 0.9).abs() < 1e-9);
        }
    }

    #[test]
    fn slide_placeholder_inherits_layout_body_properties() {
        let layout = r#"
          <p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/>
              <p:nvPr><p:ph type="title"/></p:nvPr>
            </p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="500000"/></a:xfrm></p:spPr>
            <p:txBody>
              <a:bodyPr lIns="216000" tIns="72000" rIns="216000" bIns="72000" anchor="ctr">
                <a:spAutoFit/>
              </a:bodyPr>
              <a:lstStyle/><a:p/>
            </p:txBody>
          </p:sp>"#;
        let placeholders = parse_layout_geometry(layout);
        let slide = r#"
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/>
            <p:nvPr><p:ph type="title"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>"#;

        let shape = parse_slide_shape(slide, &placeholders);
        let body = shape.text_body.expect("placeholder text body");
        assert_eq!(body.l_ins, 216_000);
        assert_eq!(body.t_ins, 72_000);
        assert_eq!(body.r_ins, 216_000);
        assert_eq!(body.b_ins, 72_000);
        assert_eq!(body.vertical_anchor, "ctr");
        assert_eq!(body.auto_fit, "sp");

        let local_left_override = slide.replace("<a:bodyPr/>", "<a:bodyPr lIns=\"0\"/>");
        let shape = parse_slide_shape(&local_left_override, &placeholders);
        let body = shape.text_body.expect("placeholder text body");
        assert_eq!(body.l_ins, 0);
        assert_eq!(body.t_ins, 72_000);
        assert_eq!(body.r_ins, 216_000);
        assert_eq!(body.b_ins, 72_000);
    }

    #[test]
    fn slide_placeholder_falls_back_to_master_text_body_properties() {
        let placeholders = LayoutPlaceholders {
            by_type_master_body_pr: HashMap::from([(
                "body".to_owned(),
                InheritedBodyPr {
                    insets: [Some(0), Some(0), Some(0), Some(0)],
                    auto_fit: Some(crate::text::InheritedAutoFit {
                        mode: "none".to_owned(),
                        font_scale: None,
                        ln_spc_reduction: None,
                    }),
                    ..InheritedBodyPr::default()
                },
            )]),
            ..LayoutPlaceholders::default()
        };
        let slide = r#"
          <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/>
            <p:nvPr><p:ph type="body" idx="10"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody>"#;

        let body = parse_slide_shape(slide, &placeholders)
            .text_body
            .expect("placeholder text body");
        assert_eq!([body.l_ins, body.t_ins, body.r_ins, body.b_ins], [0; 4]);
        assert_eq!(body.auto_fit, "none");
    }

    /// Issue #1618: every modelled bodyPr value cascades slide → layout →
    /// master → schema default, attribute by attribute, as PowerPoint renders
    /// it. The master sets each value; layout idx 11 omits them (master wins),
    /// idx 12 sets different values including explicit defaults (layout wins),
    /// and the slide's own value always wins.
    #[test]
    fn placeholder_body_properties_cascade_through_layout_and_master() {
        let master = r#"
          <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree><p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr wrap="none" vert="vert270" numCol="3" spcCol="914400" rtlCol="1"
                spcFirstLastPara="1"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>
                <a:normAutofit fontScale="50000" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/><a:p/></p:txBody>
            </p:sp></p:spTree></p:cSld>
          </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let master_body_pr = parse_master_text_body_properties(master_doc.root_element());
        let layout = r#"<p:sldLayout
              xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="11"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="12"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr wrap="square" vert="vert" numCol="2" spcCol="0" rtlCol="0"
                    spcFirstLastPara="0"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/>
                    </a:bodyPr><a:lstStyle/><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld></p:sldLayout>"#;
        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &crate::master::DefaultTextLevels::default(),
            &crate::master::MasterStyleTier::default(),
            &HashMap::new(),
            &HashMap::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &master_body_pr,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let slide = |idx: u32, body_pr: &str| {
            parse_slide_shape(
                &format!(
                    r#"<p:nvSpPr><p:cNvPr id="4" name="s"/><p:cNvSpPr/>
                      <p:nvPr><p:ph type="body" idx="{idx}"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
                    <p:txBody>{body_pr}<a:lstStyle/><a:p><a:r><a:t>Hg</a:t></a:r></a:p></p:txBody>"#
                ),
                &placeholders,
            )
            .text_body
            .expect("placeholder text body")
        };
        let summary = |b: &TextBody| {
            (
                b.wrap.clone(),
                b.vert.clone(),
                b.num_col,
                b.spc_col,
                b.rtl_col,
                b.spc_first_last_para,
                b.text_warp.as_ref().map(|w| w.preset.clone()),
                b.auto_fit.clone(),
                b.font_scale,
                b.ln_spc_reduction,
            )
        };

        let from_master = slide(11, "<a:bodyPr/>");
        assert_eq!(
            summary(&from_master),
            (
                "none".into(),
                "vert270".into(),
                3,
                914_400,
                true,
                true,
                Some("textArchUp".into()),
                "norm".into(),
                Some(0.5),
                Some(0.2)
            )
        );
        let from_layout = slide(12, "<a:bodyPr/>");
        assert_eq!(
            summary(&from_layout),
            (
                "square".into(),
                "vert".into(),
                2,
                0,
                false,
                false,
                None,
                "none".into(),
                None,
                None
            )
        );
        let slide_wins = slide(
            11,
            r#"<a:bodyPr wrap="square" vert="horz" numCol="1" spcCol="0" rtlCol="0"
                spcFirstLastPara="0"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>
                <a:normAutofit fontScale="62500"/></a:bodyPr>"#,
        );
        assert_eq!(
            summary(&slide_wins),
            (
                "square".into(),
                "horz".into(),
                1,
                0,
                false,
                false,
                None,
                "norm".into(),
                Some(0.625),
                None
            )
        );
    }

    /// Issue #1619: compatLnSpc follows the placeholder cascade, reproducing
    /// the PowerPoint control slides G1-G6. Master M1 authors compatLnSpc="1";
    /// under it layout idx 11 omits the value (G4) and idx 12 authors "0" (G5),
    /// and a slide "0" overrides the master (G6). Master M0 authors nothing;
    /// under it the value can only come from the layout (idx 13, G3) or the
    /// slide (G2), otherwise it stays unset (G1). A non-placeholder text box
    /// never takes it from the master or layout.
    /// Review probe (#1636): a bound slot whose own chain has no fontAlgn must
    /// not read a same-type sibling slot's value.
    #[test]
    fn bound_slot_font_algn_ignores_sibling_slots() {
        let layout = r#"<p:sldLayout
              xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="3"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr fontAlgn="t"/></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld></p:sldLayout>"#;
        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &DefaultTextLevels::default(),
            &MasterStyleTier::default(),
            &HashMap::new(),
            &HashMap::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        assert_eq!(placeholders.lookup_font_algn("body", Some(1)), None);
        assert_eq!(
            placeholders.lookup_font_algn("body", Some(3)),
            Some("t".to_owned())
        );
        assert_eq!(placeholders.lookup_font_algn("body", Some(7)), None);
    }

    /// ECMA-376 §21.1.2.2.7 pPr@fontAlgn follows the eaLnBrk tiers: the
    /// paragraph, the body lstStyle, the layout placeholder, the master
    /// placeholder and then the master txStyles. Only t / ctr / b survive:
    /// PowerPoint lays out omitted, auto and base identically (#1619), and an
    /// explicit base on a nearer tier still overrides an inherited t.
    #[test]
    fn font_algn_cascades_through_layout_and_master() {
        let master = r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr fontAlgn="b"/></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld>
              <p:txStyles><p:titleStyle><a:lvl1pPr fontAlgn="t"/></p:titleStyle>
                <p:bodyStyle><a:lvl1pPr fontAlgn="ctr"/></p:bodyStyle><p:otherStyle><a:lvl1pPr/></p:otherStyle></p:txStyles>
            </p:sldMaster>"#;
        let master_doc = roxmltree::Document::parse(master).unwrap();
        let styles = MasterStyleTier {
            font_algn: parse_master_font_algn_tier(master_doc.root_element(), false),
            placeholder_font_algn: parse_master_font_algn_tier(master_doc.root_element(), true),
            ..Default::default()
        };
        let layout = r#"<p:sldLayout
              xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="4" name="c"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="2"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr fontAlgn="base"/></a:lstStyle><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="5" name="d"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="3"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr fontAlgn="t"/></a:lstStyle><a:p/></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="6" name="e"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="4"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr fontAlgn="b"/></a:lstStyle><a:p/></p:txBody></p:sp>
              </p:spTree></p:cSld></p:sldLayout>"#;
        let layout_doc = roxmltree::Document::parse(layout).unwrap();
        let mut zip = empty_zip();
        let placeholders = parse_layout_placeholders(
            layout_doc.root_element(),
            &HashMap::new(),
            &DefaultTextLevels::default(),
            &styles,
            &HashMap::new(),
            &HashMap::new(),
            &MasterLevelRunProperties::default(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideLayouts/slideLayout1.xml",
            &HashMap::new(),
            &mut zip,
        );
        let effective = |nv_pr: &str, lst: &str, p_pr: &str| {
            parse_slide_shape(
                &format!(
                    r#"<p:nvSpPr><p:cNvPr id="5" name="s"/><p:cNvSpPr/>{nv_pr}</p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle>{lst}</a:lstStyle><a:p>{p_pr}<a:r><a:t>Hg</a:t></a:r></a:p></p:txBody>"#
                ),
                &placeholders,
            )
            .text_body
            .expect("text body")
            .paragraphs[0]
                .font_algn
                .clone()
        };
        let title = r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#;
        let body = r#"<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>"#;
        let pic = r#"<p:nvPr><p:ph type="pic" idx="2"/></p:nvPr>"#;
        let some = |v: &str| Some(v.to_owned());
        // The master title placeholder (b) wins over titleStyle (t).
        assert_eq!(effective(title, "", ""), some("b"));
        // bodyStyle reaches a body placeholder with no placeholder value.
        assert_eq!(effective(body, "", ""), some("ctr"));
        // A layout base overrides every master tier and is emitted as None.
        assert_eq!(effective(pic, "", ""), None);
        // The shape's own lstStyle and the paragraph win, in that order.
        assert_eq!(
            effective(body, r#"<a:lvl1pPr fontAlgn="t"/>"#, ""),
            some("t")
        );
        assert_eq!(
            effective(
                body,
                r#"<a:lvl1pPr fontAlgn="t"/>"#,
                r#"<a:pPr fontAlgn="auto"/>"#
            ),
            None
        );
        assert_eq!(
            effective(title, "", r#"<a:pPr fontAlgn="ctr"/>"#),
            some("ctr")
        );
        // An unknown token is ignored like an omitted one.
        assert_eq!(
            effective(body, "", r#"<a:pPr fontAlgn="middle"/>"#),
            some("ctr")
        );
        // Same-type body slots keep their own values (idx binding), and an
        // idx with no layout slot reads the class tier (bodyStyle).
        let slot = |idx: u32| format!(r#"<p:nvPr><p:ph type="body" idx="{idx}"/></p:nvPr>"#);
        assert_eq!(effective(&slot(3), "", ""), some("t"));
        assert_eq!(effective(&slot(4), "", ""), some("b"));
        assert_eq!(effective(&slot(9), "", ""), some("ctr"));
        // Probe: the untouched idx 1 slot keeps the master body value even
        // though its siblings (idx 3, 4) author fontAlgn.
        assert_eq!(effective(&slot(1), "", ""), some("ctr"));
        // An ordinary text box takes no placeholder tier.
        assert_eq!(effective("<p:nvPr/>", "", ""), None);
    }

    #[test]
    fn compat_ln_spc_cascades_through_layout_and_master() {
        let placeholders = |master_attr: &str| {
            let master = format!(
                r#"<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:spPr/><p:txBody><a:bodyPr{master_attr}/><a:lstStyle/><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld></p:sldMaster>"#
            );
            let master_doc = roxmltree::Document::parse(&master).unwrap();
            let master_body_pr = parse_master_text_body_properties(master_doc.root_element());
            let layout = r#"<p:sldLayout
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="11"/></p:nvPr></p:nvSpPr>
                      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>
                    <p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="12"/></p:nvPr></p:nvSpPr>
                      <p:spPr/><p:txBody><a:bodyPr compatLnSpc="0"/><a:lstStyle/><a:p/></p:txBody></p:sp>
                    <p:sp><p:nvSpPr><p:cNvPr id="4" name="c"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="13"/></p:nvPr></p:nvSpPr>
                      <p:spPr/><p:txBody><a:bodyPr compatLnSpc="1"/><a:lstStyle/><a:p/></p:txBody></p:sp>
                  </p:spTree></p:cSld></p:sldLayout>"#;
            let layout_doc = roxmltree::Document::parse(layout).unwrap();
            let mut zip = empty_zip();
            parse_layout_placeholders(
                layout_doc.root_element(),
                &HashMap::new(),
                &DefaultTextLevels::default(),
                &MasterStyleTier::default(),
                &HashMap::new(),
                &HashMap::new(),
                &MasterLevelRunProperties::default(),
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                &master_body_pr,
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                "ppt/slideLayouts/slideLayout1.xml",
                &HashMap::new(),
                &mut zip,
            )
        };
        let effective = |placeholders: &LayoutPlaceholders, nv_pr: &str, body_pr: &str| {
            parse_slide_shape(
                &format!(
                    r#"<p:nvSpPr><p:cNvPr id="5" name="s"/><p:cNvSpPr/>{nv_pr}</p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
                    <p:txBody>{body_pr}<a:lstStyle/><a:p><a:r><a:t>Hg</a:t></a:r></a:p></p:txBody>"#
                ),
                placeholders,
            )
            .text_body
            .expect("text body")
            .compat_ln_spc
        };
        let ph = |idx: u32| format!(r#"<p:nvPr><p:ph type="body" idx="{idx}"/></p:nvPr>"#);
        let text_box = "<p:nvPr/>";

        let m1 = placeholders(r#" compatLnSpc="1""#);
        assert_eq!(
            effective(&m1, &ph(11), "<a:bodyPr/>"),
            Some(true),
            "G4 master"
        );
        assert_eq!(
            effective(&m1, &ph(12), "<a:bodyPr/>"),
            Some(false),
            "G5 layout over master"
        );
        assert_eq!(
            effective(&m1, &ph(11), r#"<a:bodyPr compatLnSpc="0"/>"#),
            Some(false),
            "G6 slide over master"
        );
        assert_eq!(effective(&m1, text_box, "<a:bodyPr/>"), None, "text box");

        let m0 = placeholders("");
        assert_eq!(effective(&m0, &ph(11), "<a:bodyPr/>"), None, "G1 control");
        assert_eq!(
            effective(&m0, &ph(11), r#"<a:bodyPr compatLnSpc="1"/>"#),
            Some(true),
            "G2 slide"
        );
        assert_eq!(
            effective(&m0, &ph(13), "<a:bodyPr/>"),
            Some(true),
            "G3 layout"
        );
        assert_eq!(
            effective(&m0, &ph(12), r#"<a:bodyPr compatLnSpc="true"/>"#),
            Some(true)
        );
    }

    /// Several same-type master placeholders: each bodyPr field comes from the
    /// first placeholder in document order that sets it, so an inset on one
    /// and an autofit child on another both survive.
    #[test]
    fn master_body_properties_merge_same_type_placeholders_per_field() {
        let xml = r#"
          <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                <p:txBody><a:bodyPr lIns="0"/></p:txBody></p:sp>
              <p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
                <p:txBody><a:bodyPr lIns="12700"><a:normAutofit fontScale="50000"/></a:bodyPr></p:txBody></p:sp>
            </p:spTree></p:cSld>
          </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let master_body_pr = parse_master_text_body_properties(doc.root_element());
        let body = master_body_pr.get("body").expect("master body bodyPr");
        assert_eq!(body.insets, [Some(0), None, None, None]);
        let fit = body
            .auto_fit
            .as_ref()
            .expect("autofit from the second placeholder");
        assert_eq!((fit.mode.as_str(), fit.font_scale), ("norm", Some(0.5)));

        let placeholders = LayoutPlaceholders {
            by_type_master_body_pr: master_body_pr,
            ..LayoutPlaceholders::default()
        };
        let slide = r#"
          <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/>
            <p:nvPr><p:ph type="body" idx="10"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody>"#;
        let tb = parse_slide_shape(slide, &placeholders)
            .text_body
            .expect("placeholder text body");
        assert_eq!(tb.l_ins, 0);
        assert_eq!((tb.auto_fit.as_str(), tb.font_scale), ("norm", Some(0.5)));
    }

    #[test]
    fn master_text_body_properties_preserve_explicit_zero_insets() {
        let xml = r#"
          <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree><p:sp>
              <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
              <p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr></p:txBody>
            </p:sp></p:spTree></p:cSld>
          </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let body_pr = parse_master_text_body_properties(doc.root_element());
        let body = body_pr.get("body").expect("master body bodyPr");

        assert_eq!(body.insets, [Some(0), Some(0), Some(0), Some(0)]);
        assert_eq!(
            body.auto_fit.as_ref().map(|fit| fit.mode.as_str()),
            Some("none")
        );
    }

    #[test]
    fn slide_placeholder_inherits_vertical_anchor_from_matching_layout_idx() {
        let placeholders = parse_layout_geometry(
            r#"
              <p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Bottom body"/><p:cNvSpPr/>
                  <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
                </p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
                <p:txBody><a:bodyPr anchor="b"/><a:lstStyle/><a:p/></p:txBody>
              </p:sp>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="3" name="Top body"/><p:cNvSpPr/>
                  <p:nvPr><p:ph type="body" idx="2"/></p:nvPr>
                </p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="1000"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
                <p:txBody><a:bodyPr anchor="t"/><a:lstStyle/><a:p/></p:txBody>
              </p:sp>"#,
        );
        let shape = parse_slide_shape(
            r#"
              <p:nvSpPr><p:cNvPr id="4" name="Top body instance"/><p:cNvSpPr/>
                <p:nvPr><p:ph type="body" idx="2"/></p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Top</a:t></a:r></a:p></p:txBody>"#,
            &placeholders,
        );

        assert_eq!(
            shape
                .text_body
                .expect("placeholder text body")
                .vertical_anchor,
            "t",
        );

        let master_only = LayoutPlaceholders {
            by_type_master_anchor: HashMap::from([("".to_owned(), "b".to_owned())]),
            ..LayoutPlaceholders::default()
        };
        assert_eq!(
            master_only.lookup_anchor("obj", Some(99)).as_deref(),
            Some("b"),
        );
    }

    #[test]
    fn idx_placeholder_without_anchor_retains_layout_type_fallback() {
        let placeholders = parse_layout_geometry(
            r#"
              <p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Shared body anchor"/><p:cNvSpPr/>
                  <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>
              </p:sp>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="3" name="Body instance"/><p:cNvSpPr/>
                  <p:nvPr><p:ph type="body" idx="10"/></p:nvPr>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
              </p:sp>"#,
        );

        assert_eq!(
            placeholders.lookup_anchor("body", Some(10)).as_deref(),
            Some("ctr"),
        );
    }

    const LAYOUT_ELLIPSE: &str = r#"
        <p:sp>
          <p:nvSpPr><p:cNvPr id="27" name="Quarter"/><p:cNvSpPr/>
            <p:nvPr><p:ph type="body" idx="18"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>
            <a:prstGeom prst="ellipse"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
          </p:spPr>
        </p:sp>"#;

    const SLIDE_PLACEHOLDER: &str = r#"
        <p:nvSpPr><p:cNvPr id="27" name="Quarter"/><p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="18"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>
        </p:spPr>"#;

    #[test]
    fn slide_placeholder_inherits_geometry_and_adjustments_from_matching_layout_idx() {
        let placeholders = parse_layout_geometry(LAYOUT_ELLIPSE);
        let shape = parse_slide_shape(SLIDE_PLACEHOLDER, &placeholders);

        assert_eq!(shape.geometry, "ellipse");
        assert_eq!(shape.adj, Some(25000.0));
        assert_eq!(
            (shape.x, shape.y, shape.width, shape.height),
            (100, 200, 300, 400)
        );
    }

    #[test]
    fn slide_placeholder_local_geometry_overrides_layout_geometry() {
        let placeholders = parse_layout_geometry(LAYOUT_ELLIPSE);
        let own_rect = SLIDE_PLACEHOLDER.replace(
            "</p:spPr>",
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>",
        );
        let shape = parse_slide_shape(&own_rect, &placeholders);

        assert_eq!(shape.geometry, "rect");
        assert_eq!(shape.adj, None);
    }

    #[test]
    fn explicit_idx_does_not_borrow_geometry_from_another_layout_slot() {
        let placeholders = parse_layout_geometry(LAYOUT_ELLIPSE);
        let different_idx = SLIDE_PLACEHOLDER.replace("idx=\"18\"", "idx=\"19\"");
        let shape = parse_slide_shape(&different_idx, &placeholders);

        assert_eq!(shape.geometry, "rect");
    }
}
