//! Slide-master / layout inheritance: the per-master extractors (anchors,
//! alignments, ea-line-break, font sizes, per-level sizes/indents/bullets,
//! txStyle bold/italic/colour/spacing, transforms), the layout-placeholder
//! resolver, and the cached `ParsedMaster` (was `MasterBundle`) / `ParsedLayout`
//! bundles. Extracted verbatim from `lib.rs`; the only non-move change is the
//! `MasterBundle` → `ParsedMaster` type rename (fields unchanged).

use crate::fill::{
    parse_background, parse_blip_fill, parse_color_node, parse_cust_geom, parse_fill,
    parse_reflection, parse_xfrm,
};
use crate::shape::{
    extract_decorative_shapes, resolve_picture_shape_properties, PictureShapeProperties,
};
use crate::text::{
    empty_level_bullets, extract_level_bullets, extract_level_font_sizes, extract_level_indents,
    extract_lvl1_font_size, has_any_level_bullet, has_any_level_indent, has_any_level_size,
    merge_level_bullets, merge_level_indents, merge_level_sizes, read_level_bullets,
    read_level_font_sizes, read_level_indents, text_property_solid_fill, BuMarker, LevelBullets,
    LevelFontSizes, LevelIndents,
};
use crate::theme::{
    bake_clr_map, parse_theme_part, resolve_theme_typeface, PptxTheme,
    PptxThemeSource,
};
use crate::types::*;
use crate::{
    attr, attr_f64, attr_i64, attr_r, build_smartart_drawings, child, find_rel_target_by_type,
    note_layout_master_parse, parse_preflighted_pptx_xml, parse_rels, read_zip_str, resolve_path,
    PptxZip,
};
use ooxml_common::blip::{mime_from_ext, Duotone, SrcRect};
use ooxml_common::rels::relationship_part_path;
use std::collections::HashMap;

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
    /// Default font size (pt) per placeholder idx, from layout/master lstStyle
    pub(crate) by_idx_font_size: HashMap<u32, f64>,
    /// Default font size (pt) per placeholder type, from layout/master lstStyle
    pub(crate) by_type_font_size: HashMap<String, f64>,
    /// Master-only font sizes retained separately so an idx-bearing slide
    /// placeholder that has no matching layout slot can still inherit the
    /// master without borrowing an unrelated layout sibling of the same type.
    pub(crate) by_type_master_font_size: HashMap<String, f64>,
    /// Default Latin typeface per placeholder idx/type, inherited from the
    /// layout placeholder's lstStyle and then the master txStyles. Theme font
    /// tokens such as +mj-lt are resolved before storage.
    pub(crate) by_idx_font_family: HashMap<u32, String>,
    pub(crate) by_type_font_family: HashMap<String, String>,
    pub(crate) by_type_master_font_family: HashMap<String, String>,
    /// Per-list-level default font sizes (pt) per placeholder idx — index 0..=8
    /// maps to lvl1pPr..lvl9pPr (ECMA-376 §21.1.2.4). Lets nested bullets shrink
    /// per level (e.g. body 28pt → lvl2 24pt → lvl3 20pt) instead of all using
    /// the level-1 size. None per level where the style chain doesn't specify it.
    pub(crate) by_idx_level_sizes: HashMap<u32, LevelFontSizes>,
    /// Per-list-level default font sizes (pt) per placeholder type.
    pub(crate) by_type_level_sizes: HashMap<String, LevelFontSizes>,
    pub(crate) by_type_master_level_sizes: HashMap<String, LevelFontSizes>,
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
    /// Per-placeholder layout `bodyPr` text insets (`lIns`, `tIns`, `rIns`,
    /// `bIns`). Each component stays optional so an omitted layout attribute
    /// can continue through the theme/spec fallback instead of being replaced
    /// by a synthetic layout default.
    pub(crate) by_idx_text_insets: HashMap<u32, [Option<i64>; 4]>,
    pub(crate) by_type_text_insets: HashMap<String, [Option<i64>; 4]>,
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
    /// Default space-before (hundredths of pt) per placeholder type, from layout lstStyle
    pub(crate) by_type_space_before: HashMap<String, i64>,
    /// Default space-after (hundredths of pt) per placeholder type, from layout lstStyle
    pub(crate) by_type_space_after: HashMap<String, i64>,
    /// Default space-before from master txStyles (fallback when layout has none)
    pub(crate) by_type_master_space_before: HashMap<String, i64>,
    /// Default space-after from master txStyles (fallback when layout has none)
    pub(crate) by_type_master_space_after: HashMap<String, i64>,
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
    /// Default line spacing (spcPct val, e.g. 90000 = 90%) per placeholder idx, from layout lstStyle
    pub(crate) by_idx_line_spacing: HashMap<u32, f64>,
    /// Default line spacing (spcPct val) per placeholder type, from layout lstStyle
    pub(crate) by_type_line_spacing: HashMap<String, f64>,
    /// Paragraph alignment per placeholder type from master lstStyle > lvl1pPr algn (fallback)
    pub(crate) by_type_master_alignment: HashMap<String, String>,
    /// East Asian line-break per placeholder type from master lstStyle > lvl1pPr
    /// @eaLnBrk (fallback when the layout has none) — ECMA-376 §21.1.2.2.7
    pub(crate) by_type_master_ea_ln_brk: HashMap<String, bool>,
    /// Default line spacing from master txStyles (fallback when layout has none)
    pub(crate) by_type_master_line_spacing: HashMap<String, f64>,
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
            return Some(Self {
                geometry: "custGeom".to_owned(),
                cust_geom: Some(parse_cust_geom(cust_geom_node, shape_w, shape_h)),
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
        let adjustment = |index: usize| -> Option<f64> {
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
        };

        Some(Self {
            geometry,
            cust_geom: None,
            adjustments: std::array::from_fn(adjustment),
        })
    }
}

#[derive(Debug, Clone, serde::Serialize)]
pub(crate) struct InheritedBlipFill {
    /// Embedded zip path of the inherited picture-placeholder blip.
    pub(crate) image_path: String,
    /// MIME of the blip at `image_path`.
    pub(crate) mime_type: String,
    pub(crate) svg_image_path: Option<String>,
    pub(crate) dpi: Option<u32>,
    pub(crate) rot_with_shape: Option<bool>,
    pub(crate) src_rect: Option<SrcRect>,
    pub(crate) fill_rect: Option<FillRect>,
    pub(crate) tile: Option<TileInfo>,
    pub(crate) stretch: bool,
    pub(crate) alpha: Option<f64>,
    /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour on the layout placeholder's
    /// blipFill, resolved through the theme. Inherited onto the slide picture
    /// placeholder that omits its own blipFill (see `shape.rs`).
    pub(crate) duotone: Option<Duotone>,
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

    /// Look up the inherited default font size for a placeholder (layout then master fallback).
    /// Idx-strict per ECMA-376 §19.3.1.36 (see `lookup_fill`'s rationale).
    pub(crate) fn lookup_font_size(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        if let Some(i) = ph_idx {
            return self.by_idx_font_size.get(&i).copied().or_else(|| {
                self.by_type_master_font_size
                    .get(ph_type)
                    .copied()
                    .or_else(|| {
                        if ph_type == "obj" {
                            self.by_type_master_font_size.get("").copied()
                        } else {
                            None
                        }
                    })
            });
        }
        self.by_type_font_size.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_font_size.get("").copied()
            } else {
                None
            }
        })
    }

    pub(crate) fn lookup_font_family(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
        if let Some(i) = ph_idx {
            return self.by_idx_font_family.get(&i).cloned().or_else(|| {
                self.by_type_master_font_family
                    .get(ph_type)
                    .cloned()
                    .or_else(|| {
                        if ph_type == "obj" {
                            self.by_type_master_font_family.get("").cloned()
                        } else {
                            None
                        }
                    })
            });
        }
        self.by_type_font_family.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_font_family.get("").cloned()
            } else {
                None
            }
        })
    }

    /// Per-list-level inherited default font sizes (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_font_size`. All-None when the placeholder has no
    /// per-level styling.
    pub(crate) fn lookup_level_font_sizes(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> LevelFontSizes {
        if let Some(i) = ph_idx {
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

    /// Per-list-level inherited paragraph indents (lvl1..lvl9). Same idx-strict
    /// resolution as `lookup_level_font_sizes`. All-default (every axis None) when
    /// the placeholder has no authored per-level indent.
    pub(crate) fn lookup_level_indents(&self, ph_type: &str, ph_idx: Option<u32>) -> LevelIndents {
        if let Some(i) = ph_idx {
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
    pub(crate) fn lookup_bold(&self, ph_type: &str) -> Option<bool> {
        self.by_type_bold.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_bold.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited italic for this placeholder type.
    pub(crate) fn lookup_italic(&self, ph_type: &str) -> Option<bool> {
        self.by_type_italic.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_italic.get("").copied()
            } else {
                None
            }
        })
    }

    /// Look up inherited caps ("all"/"small") for this placeholder type.
    pub(crate) fn lookup_caps(&self, ph_type: &str) -> Option<String> {
        self.by_type_caps.get(ph_type).cloned().or_else(|| {
            if ph_type == "body" {
                self.by_type_caps.get("").cloned()
            } else {
                None
            }
        })
    }

    pub(crate) fn lookup_reflection(&self, ph_type: &str) -> Option<Reflection> {
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

    /// Look up layout placeholder text insets. An explicit idx is strict so a
    /// body placeholder cannot borrow another body slot's margins.
    pub(crate) fn lookup_text_insets(
        &self,
        ph_type: &str,
        ph_idx: Option<u32>,
    ) -> Option<[Option<i64>; 4]> {
        if let Some(i) = ph_idx {
            return self.by_idx_text_insets.get(&i).copied();
        }
        self.by_type_text_insets.get(ph_type).copied().or_else(|| {
            if ph_type == "body" {
                self.by_type_text_insets.get("").copied()
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
    pub(crate) fn lookup_ea_ln_brk(&self, ph_type: &str) -> Option<bool> {
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

    pub(crate) fn lookup_space_before(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_before
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_space_before.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_space_before.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_space_before.get("").copied()
                } else {
                    None
                }
            })
    }

    pub(crate) fn lookup_space_after(&self, ph_type: &str) -> Option<i64> {
        self.by_type_space_after
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_space_after.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_space_after.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_space_after.get("").copied()
                } else {
                    None
                }
            })
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
    /// The *master* `txStyles` tier (titleStyle/bodyStyle/otherStyle), however, is a
    /// document-wide default keyed by placeholder *type* (§21.1.2.4 / §19.3.1) and is
    /// inherited regardless of idx. So when the idx-matched layout shape defines no
    /// colour, resolution must still fall through to `by_type_master_color`. Without
    /// this, a body placeholder whose layout shape sets size-but-not-colour resolves to
    /// no colour at all and the renderer defaults to black — instead of the master
    /// bodyStyle colour (e.g. `schemeClr bg1` = white on a dark theme). (sample-9 slide 2+)
    pub(crate) fn lookup_color(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<String> {
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

    /// Look up inherited line spacing (spcPct val, e.g. 90000 = 90%) for this placeholder.
    /// Idx-strict per ECMA-376 §19.3.1.36 (see `lookup_fill`'s rationale).
    pub(crate) fn lookup_line_spacing(&self, ph_type: &str, ph_idx: Option<u32>) -> Option<f64> {
        if let Some(i) = ph_idx {
            return self.by_idx_line_spacing.get(&i).copied();
        }
        self.by_type_line_spacing
            .get(ph_type)
            .copied()
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_line_spacing.get("").copied()
                } else {
                    None
                }
            })
            .or_else(|| self.by_type_master_line_spacing.get(ph_type).copied())
            .or_else(|| {
                if ph_type == "body" {
                    self.by_type_master_line_spacing.get("").copied()
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

/// txStyles style node → the placeholder types it defaults. ECMA-376 §19.3.1.52
/// txStyles → titleStyle §19.3.1.49 / bodyStyle §19.3.1.5 / otherStyle §19.3.1.35.
pub(crate) const MASTER_TXSTYLE_PH_TYPES: &[(&str, &[&str])] = &[
    ("titleStyle", &["title", "ctrTitle"]),
    ("bodyStyle", &["body", "subTitle", "obj", ""]),
    ("otherStyle", &["dt", "ftr", "sldNum"]),
];

/// Parse paragraph alignment from master placeholder shapes' lstStyle > lvl1pPr algn attribute.
/// Takes the shared, already-parsed master root (see `parse_master_anchors`).
pub(crate) fn parse_master_alignments(root: roxmltree::Node<'_, '_>) -> HashMap<String, String> {
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
                if let Some(algn) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| attr(&lp, "algn"))
                {
                    map.entry(ph_type).or_insert(algn);
                }
            }
        }
    }
    // Fallback: master <p:txStyles> paragraph alignment (ECMA-376 §19.3.1.52
    // txStyles → titleStyle §19.3.1.49 / bodyStyle §19.3.1.5 / otherStyle §19.3.1.35).
    // Per-shape lstStyle (scanned above) wins via or_insert; this fills types
    // whose master placeholder shape carried no explicit algn (the common case —
    // PowerPoint stores title/body alignment in txStyles, not the shape lstStyle).
    if let Some(tx_styles) = child(root, "txStyles") {
        for &(style, types) in MASTER_TXSTYLE_PH_TYPES {
            if let Some(algn) = child(tx_styles, style)
                .and_then(|s| child(s, "lvl1pPr"))
                .and_then(|lp| attr(&lp, "algn"))
            {
                for t in types {
                    map.entry((*t).to_string()).or_insert_with(|| algn.clone());
                }
            }
        }
    }
    map
}

/// Parse master-level default East Asian line-break (eaLnBrk) per placeholder
/// type from each placeholder shape's lstStyle > lvl1pPr @eaLnBrk
/// (ECMA-376 §21.1.2.2.7). Mirrors parse_master_alignments. xsd:boolean.
pub(crate) fn parse_master_ea_ln_brk(root: roxmltree::Node<'_, '_>) -> HashMap<String, bool> {
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
                if let Some(v) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| attr(&lp, "eaLnBrk"))
                {
                    map.entry(ph_type).or_insert(v == "1" || v == "true");
                }
            }
        }
    }
    map
}

/// Parse master-level default font sizes from txStyles (titleStyle / bodyStyle / otherStyle)
/// and from individual placeholder shapes in the master spTree.
/// Individual shape lstStyle takes priority over txStyles generic defaults.
pub(crate) fn parse_master_font_sizes(root: roxmltree::Node<'_, '_>) -> HashMap<String, f64> {
    let mut map = HashMap::new();

    // Scan master spTree placeholder shapes first — per-shape lstStyle is more specific
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
                if let Some(tx_body) = child(sp, "txBody") {
                    if let Some(sz) = extract_lvl1_font_size(tx_body) {
                        map.entry(ph_type).or_insert(sz);
                    }
                }
            }
        }
    }

    // p:txStyles > a:titleStyle / a:bodyStyle / a:otherStyle as fallback
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            let sz = child(tx_styles, style_name)
                .and_then(|sn| child(sn, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(|rp| attr_f64(&rp, "sz"))
                .map(|v| v / 100.0);
            if let Some(fs) = sz {
                for ph_type in *ph_types {
                    map.entry(ph_type.to_string()).or_insert(fs);
                }
            }
        }
    } else {
        // CT_SlideMaster permits txStyles to be omitted. In that case current
        // PowerPoint supplies its application-level placeholder defaults:
        // titles at 44 pt, body/subtitle placeholders at 28 pt, and object
        // placeholders at 18 pt. Scope these
        // compatibility defaults to presentation placeholders only; ordinary
        // text boxes continue through their own authored/theme cascade.
        for ph_type in ["title", "ctrTitle"] {
            map.entry(ph_type.to_owned()).or_insert(44.0);
        }
        for ph_type in ["body", "subTitle"] {
            map.entry(ph_type.to_owned()).or_insert(28.0);
        }
        for ph_type in ["obj", ""] {
            map.entry(ph_type.to_owned()).or_insert(18.0);
        }
    }

    map
}

/// Default Latin typefaces from master placeholder lstStyle/txStyles. The
/// specific placeholder shape wins over the generic title/body/other style,
/// matching the font-size cascade above (ECMA-376 §19.3.1.52 and
/// §21.1.2.3.7). Theme tokens are resolved against this master's theme.
pub(crate) fn parse_master_font_families(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let read_family = |def_rpr: roxmltree::Node<'_, '_>| {
        child(def_rpr, "latin")
            .and_then(|latin| attr(&latin, "typeface"))
            .map(|face| resolve_theme_typeface(&face, theme))
    };

    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(family) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| child(lp, "defRPr"))
                    .and_then(read_family)
                {
                    map.entry(ph_type).or_insert(family);
                }
            }
        }
    }

    if let Some(tx_styles) = child(root, "txStyles") {
        for &(style_name, ph_types) in MASTER_TXSTYLE_PH_TYPES {
            if let Some(family) = child(tx_styles, style_name)
                .and_then(|style| child(style, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(read_family)
            {
                for ph_type in ph_types {
                    map.entry((*ph_type).to_owned())
                        .or_insert_with(|| family.clone());
                }
            }
        }
    }
    map
}

/// Per-list-level default font sizes from the master, keyed by ph_type. Mirrors
/// `parse_master_font_sizes` but captures every list level (lvl1pPr..lvl9pPr) so
/// nested bullets inherit the correct shrinking sizes (ECMA-376 §21.1.2.4),
/// not just the level-1 size. Per-shape lstStyle wins over the generic txStyles.
pub(crate) fn parse_master_level_font_sizes(
    root: roxmltree::Node<'_, '_>,
) -> HashMap<String, LevelFontSizes> {
    let mut map: HashMap<String, LevelFontSizes> = HashMap::new();

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let sizes = extract_level_font_sizes(tx_body);
                    if has_any_level_size(&sizes) {
                        map.entry(ph_type).or_insert(sizes);
                    }
                }
            }
        }
    }

    // txStyles fallback.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let sizes = read_level_font_sizes(style_node);
                if has_any_level_size(&sizes) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string()).or_insert(sizes);
                    }
                }
            }
        }
    }

    map
}

/// Per-list-level paragraph indents (`marL`/`marR`/`indent`, EMU) from the master,
/// keyed by ph_type. Mirrors `parse_master_level_font_sizes` exactly (same per-shape
/// lstStyle then `txStyles` tiers via `MASTER_TXSTYLE_PH_TYPES`): a master body
/// `<a:lvlNpPr@marL>` is what a slide body paragraph with no own `marL` inherits
/// (ECMA-376 §21.1.2.4.13). KNOWN SHARED GAP: no presentation `defaultTextStyle`
/// tier (§19.2.1.8, the lowest authored fallback) — the parser reads it for neither
/// indents nor font sizes nor bullets, so this stays at parity rather than adding a
/// tier only here; closing it is a separate cross-cutting change.
pub(crate) fn parse_master_level_indents(
    root: roxmltree::Node<'_, '_>,
) -> HashMap<String, LevelIndents> {
    let mut map: HashMap<String, LevelIndents> = HashMap::new();

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let indents = extract_level_indents(tx_body);
                    if has_any_level_indent(&indents) {
                        map.entry(ph_type).or_insert(indents);
                    }
                }
            }
        }
    }

    // txStyles fallback.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let indents = read_level_indents(style_node);
                if has_any_level_indent(&indents) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string()).or_insert(indents);
                    }
                }
            }
        }
    } else {
        // Office application defaults observed with an omitted txStyles tier:
        // level-1 body/subtitle paragraphs use an 18 pt hanging gutter. Only
        // this observed boundary is synthesized; object/content placeholders
        // and deeper levels remain unstyled.
        let mut level_one: LevelIndents = Default::default();
        level_one[0].mar_l = Some(228_600);
        level_one[0].indent = Some(-228_600);
        for ph_type in ["body", "subTitle"] {
            map.entry(ph_type.to_owned()).or_insert(level_one);
        }
    }

    map
}

/// Per-list-level bullets from the master, keyed by ph_type. Mirrors
/// `parse_master_level_font_sizes`: a master body placeholder's `<a:buChar>` (or
/// the `bodyStyle` `<a:lvlNpPr>` bullets) is what a slide body paragraph with no
/// explicit bullet inherits (ECMA-376 §19.7.10 / §21.1.2.4). Per-shape lstStyle
/// wins over the generic txStyles.
pub(crate) fn parse_master_level_bullets(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
    master_rels: &HashMap<String, String>,
    master_dir: &str,
    zip: &mut PptxZip,
) -> HashMap<String, LevelBullets> {
    let mut map: HashMap<String, LevelBullets> = HashMap::new();

    // A master-level `<a:buBlip>` embed resolves against the master's rels +
    // part directory (ECMA-376 §21.1.2.4.2), mirroring the master background.
    let mut resolve_blip = |rid: &str| -> Option<String> {
        let target = master_rels.get(rid)?;
        let path = resolve_path(master_dir, target);
        // Verify the part exists so a listed-but-missing rId yields None and the
        // bullet falls through to Bullet::Inherit (matches the variant's doc
        // comment), mirroring the master background resolver. `index_for_name`
        // checks the central directory only (no inflate), unlike the former
        // `read_zip_bytes` which decompressed the entry just to discard it.
        zip.index_for_name(&path)?;
        Some(path)
    };

    // Per-shape lstStyle first (more specific).
    if let Some(sp_tree) = child(root, "cSld").and_then(|n| child(n, "spTree")) {
        for sp in sp_tree
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "sp")
        {
            if let Some(ph) = sp
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "ph")
            {
                let ph_type = attr(&ph, "type").unwrap_or_default();
                if let Some(tx_body) = child(sp, "txBody") {
                    let bullets = extract_level_bullets(tx_body, theme, &mut resolve_blip);
                    if has_any_level_bullet(&bullets) {
                        map.entry(ph_type).or_insert(bullets);
                    }
                }
            }
        }
    }

    // txStyles fallback. Within the master tier the per-shape placeholder
    // `lstStyle` (above) is MORE specific than the generic `txStyles`, but the
    // two are still resolved per-group: a per-shape entry that declares only a
    // marker inherits its colour/size/font from the matching `txStyles` level
    // (ECMA-376 §21.1.2.4 — the four bullet groups inherit independently). So
    // merge the existing per-shape entry (primary) over the txStyles bullets
    // (fallback) rather than dropping txStyles wholesale.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(style_node) = child(tx_styles, style_name) {
                let bullets = read_level_bullets(style_node, theme, &mut resolve_blip);
                if has_any_level_bullet(&bullets) {
                    for ph_type in *ph_types {
                        map.entry(ph_type.to_string())
                            .and_modify(|existing| {
                                *existing = merge_level_bullets(existing, &bullets)
                            })
                            .or_insert_with(|| bullets.clone());
                    }
                }
            }
        }
    } else {
        // Office application defaults observed with an omitted txStyles tier:
        // level-1 body/subtitle paragraphs use a round bullet. Keep every
        // decoration group inherited so marker colour/size/typeface follows
        // the resolved text, and do not extrapolate to object/deeper levels.
        let mut level_one = empty_level_bullets();
        level_one[0].marker = Some(BuMarker::Char("•".to_owned()));
        for ph_type in ["body", "subTitle"] {
            map.entry(ph_type.to_owned())
                .or_insert_with(|| level_one.clone());
        }
    }

    map
}

/// Parse default bold/italic from master txStyles (titleStyle / bodyStyle / otherStyle)
/// > lvl1pPr > defRPr @b and @i. Keyed by ph_type.
/// > Only populated when the attribute is explicitly present on the master.
type MasterTxStyleRunProperties = (
    HashMap<String, bool>,
    HashMap<String, bool>,
    HashMap<String, String>,
    HashMap<String, Reflection>,
);

pub(crate) fn parse_master_txstyle_run_properties(
    root: roxmltree::Node<'_, '_>,
) -> MasterTxStyleRunProperties {
    let mut bold_map: HashMap<String, bool> = HashMap::new();
    let mut italic_map: HashMap<String, bool> = HashMap::new();
    // ECMA-376 §21.1.2.3.9, ST_TextCapsType §20.1.10.64: cap="all"/"small"
    // on the master txStyles defRPr —
    // e.g. a template titleStyle with cap="all" upper-cases every title.
    let mut caps_map: HashMap<String, String> = HashMap::new();
    let mut reflection_map: HashMap<String, Reflection> = HashMap::new();
    let Some(tx_styles) = child(root, "txStyles") else {
        return (bold_map, italic_map, caps_map, reflection_map);
    };
    let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
    for (style_name, ph_types) in style_ph_map {
        let def_rpr = child(tx_styles, style_name)
            .and_then(|sn| child(sn, "lvl1pPr"))
            .and_then(|lp| child(lp, "defRPr"));
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
        if let Some(bv) = b {
            for t in *ph_types {
                bold_map.entry(t.to_string()).or_insert(bv);
            }
        }
        if let Some(iv) = i {
            for t in *ph_types {
                italic_map.entry(t.to_string()).or_insert(iv);
            }
        }
        if let Some(cv) = c {
            for t in *ph_types {
                caps_map.entry(t.to_string()).or_insert(cv.clone());
            }
        }
        if let Some(value) = reflection {
            for t in *ph_types {
                reflection_map
                    .entry(t.to_string())
                    .or_insert_with(|| value.clone());
            }
        }
    }
    (bold_map, italic_map, caps_map, reflection_map)
}

/// Parse default text color from master txStyles (titleStyle/bodyStyle/otherStyle)
/// > lvl1pPr > defRPr > solidFill, and from per-placeholder shapes in the master spTree's
/// > txBody > lstStyle > lvl1pPr > defRPr > solidFill. Keyed by ph_type.
/// > Shape-level lstStyle takes priority over txStyles generic defaults.
pub(crate) fn parse_master_txstyle_color(
    root: roxmltree::Node<'_, '_>,
    theme: &HashMap<String, String>,
) -> HashMap<String, String> {
    let mut map: HashMap<String, String> = HashMap::new();

    // Scan master spTree placeholder shapes first — per-shape lstStyle is more specific.
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
                if let Some(color) = child(sp, "txBody")
                    .and_then(|tb| child(tb, "lstStyle"))
                    .and_then(|ls| child(ls, "lvl1pPr"))
                    .and_then(|lp| child(lp, "defRPr"))
                    .and_then(text_property_solid_fill)
                    .and_then(|sf| parse_color_node(sf, theme))
                {
                    map.entry(ph_type).or_insert(color);
                }
            }
        }
    }

    // Fall back to p:txStyles > titleStyle/bodyStyle/otherStyle > lvl1pPr > defRPr > solidFill.
    if let Some(tx_styles) = child(root, "txStyles") {
        let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
        for (style_name, ph_types) in style_ph_map {
            if let Some(color) = child(tx_styles, style_name)
                .and_then(|sn| child(sn, "lvl1pPr"))
                .and_then(|lp| child(lp, "defRPr"))
                .and_then(text_property_solid_fill)
                .and_then(|sf| parse_color_node(sf, theme))
            {
                for ph_type in *ph_types {
                    map.entry(ph_type.to_string()).or_insert(color.clone());
                }
            }
        }
    }

    map
}

/// Parse default paragraph spacing from master txStyles.
/// Returns (space_before_map, space_after_map, line_spacing_map) keyed by ph_type string.
/// space_before/after values are in hundredths of a point (same as Paragraph.space_before/after).
/// Note: line_spacing_map is intentionally NOT populated. Inheriting txStyles lnSpc hurts VRT
/// scores because our font substitutes (sans-serif) have different em-square metrics than the
/// original Aptos font, so applying the master's 120% line spacing over-expands text layout.
pub(crate) fn parse_master_txstyle_spacing(
    root: roxmltree::Node<'_, '_>,
) -> (
    HashMap<String, i64>,
    HashMap<String, i64>,
    HashMap<String, f64>,
) {
    let mut before_map: HashMap<String, i64> = HashMap::new();
    let mut after_map: HashMap<String, i64> = HashMap::new();
    let line_map: HashMap<String, f64> = HashMap::new(); // intentionally not populated
    let tx_styles = match child(root, "txStyles") {
        Some(n) => n,
        None => return (before_map, after_map, line_map),
    };
    let style_ph_map: &[(&str, &[&str])] = MASTER_TXSTYLE_PH_TYPES;
    for (style_name, ph_types) in style_ph_map {
        let lvl1 = child(tx_styles, style_name).and_then(|sn| child(sn, "lvl1pPr"));
        let spc_before = lvl1
            .and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        let spc_after = lvl1
            .and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts").and_then(|n| attr_i64(&n, "val")));
        if let Some(v) = spc_before {
            for ph_type in *ph_types {
                before_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
        if let Some(v) = spc_after {
            for ph_type in *ph_types {
                after_map.entry(ph_type.to_string()).or_insert(v);
            }
        }
    }
    (before_map, after_map, line_map)
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
    master_font_sizes: &HashMap<String, f64>,
    master_font_families: &HashMap<String, String>,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_space_before: &HashMap<String, i64>,
    master_space_after: &HashMap<String, i64>,
    master_line_spacing: &HashMap<String, f64>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    layout_dir: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> LayoutPlaceholders {
    let theme = theme_source.colors();
    let mut lph = LayoutPlaceholders {
        master_by_type: master_transforms.clone(),
        by_type_master_font_size: master_font_sizes.clone(),
        by_type_master_font_family: master_font_families.clone(),
        by_type_master_level_sizes: master_level_font_sizes.clone(),
        by_type_master_level_indents: master_level_indents.clone(),
        by_type_master_level_bullets: master_level_bullets.clone(),
        by_type_master_anchor: master_anchors.clone(),
        by_type_master_alignment: master_alignments.clone(),
        by_type_master_ea_ln_brk: master_ea_ln_brk.clone(),
        by_type_master_space_before: master_space_before.clone(),
        by_type_master_space_after: master_space_after.clone(),
        by_type_master_line_spacing: master_line_spacing.clone(),
        ..Default::default()
    };

    let sp_tree = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "spTree");
    let sp_tree = match sp_tree {
        Some(n) => n,
        None => return lph,
    };

    for sp in sp_tree
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "sp")
    {
        let ph_node = sp
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "ph");
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
        let layout_font_size = layout_def_rpr
            .and_then(|rp| attr_f64(&rp, "sz"))
            .map(|v| v / 100.0);
        let layout_font_family = layout_def_rpr
            .and_then(|rp| child(rp, "latin"))
            .and_then(|latin| attr(&latin, "typeface"))
            .map(|face| resolve_theme_typeface(&face, theme));
        // Per-level sizes from the layout placeholder's own lstStyle (all
        // lvlNpPr), used to give nested bullets their shrinking sizes.
        let layout_level_sizes: LevelFontSizes = child(sp, "txBody")
            .map(extract_level_font_sizes)
            .unwrap_or([None; 9]);
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
            let path = resolve_path(layout_dir, target);
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
        let layout_color: Option<String> = layout_def_rpr
            .and_then(text_property_solid_fill)
            .and_then(|sf| parse_color_node(sf, theme));
        let layout_alignment: Option<String> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "algn"))
            .map(|a| a.to_string());
        // ECMA-376 §21.1.2.2.7 eaLnBrk from the layout placeholder's lvl1pPr.
        let layout_ea_ln_brk: Option<bool> = layout_lvl1_ppr
            .and_then(|lp| attr(&lp, "eaLnBrk"))
            .map(|v| v == "1" || v == "true");
        let layout_space_before: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcBef"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        let layout_space_after: Option<i64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "spcAft"))
            .and_then(|s| child(s, "spcPts"))
            .and_then(|s| attr_i64(&s, "val"));
        // lnSpc > spcPct val (e.g. 90000 = 90%)
        let layout_line_spacing: Option<f64> = layout_lvl1_ppr
            .and_then(|lp| child(lp, "lnSpc"))
            .and_then(|ls| child(ls, "spcPct"))
            .and_then(|s| attr_f64(&s, "val"));

        let layout_body_pr = child(sp, "txBody").and_then(|tb| child(tb, "bodyPr"));
        // Layout bodyPr anchor; fall back to master anchor map.
        let layout_anchor: Option<String> = layout_body_pr
            .and_then(|bp| attr(&bp, "anchor"))
            .map(|a| a.to_string());
        let layout_text_insets: [Option<i64>; 4] = [
            layout_body_pr.and_then(|bp| attr_i64(&bp, "lIns")),
            layout_body_pr.and_then(|bp| attr_i64(&bp, "tIns")),
            layout_body_pr.and_then(|bp| attr_i64(&bp, "rIns")),
            layout_body_pr.and_then(|bp| attr_i64(&bp, "bIns")),
        ];
        let has_layout_text_inset = layout_text_insets.iter().any(Option::is_some);

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
            let image_path = resolve_path(layout_dir, rel_target);
            // Verify the part exists so a dangling rId yields None (no inherited
            // fill), preserving the prior data-URL behaviour. `index_for_name`
            // reads the central directory only (no inflate), unlike the former
            // `read_zip_bytes` which decompressed the entry just to discard it.
            zip.index_for_name(&image_path)?;
            let mime_type = mime_from_ext(&image_path).to_owned();
            let mut resolve = |relationship_id: &str| {
                let target = layout_rels.get(relationship_id)?;
                let path = resolve_path(layout_dir, target);
                zip.index_for_name(&path)?;
                Some(path)
            };
            let Fill::Image {
                svg_image_path, dpi, rot_with_shape, src_rect, fill_rect,
                stretch, tile, alpha, duotone, ..
            } = parse_blip_fill(bf, theme, &mut resolve)? else {
                return None;
            };
            Some(InheritedBlipFill { image_path, mime_type, svg_image_path, dpi,
                rot_with_shape, src_rect, fill_rect, tile, stretch, alpha, duotone })
        });

        if let Some(ph) = ph_node {
            // CT_Placeholder defaults an omitted @type to `obj`. The idx binds
            // the slide placeholder to this layout slot, but does not permit a
            // same-numbered master placeholder of another type to rewrite the
            // schema value. Real PowerPoint layouts can reuse an idx for a
            // content slot where the master uses it for date/footer metadata.
            let ph_idx: Option<u32> = attr(&ph, "idx").and_then(|v| v.parse().ok());
            let ph_type = attr(&ph, "type").unwrap_or_else(|| "obj".to_owned());

            if let Some(idx) = ph_idx {
                lph.by_idx_placeholder_type
                    .entry(idx)
                    .or_insert_with(|| ph_type.clone());
                if let Some(ref t) = t_opt {
                    lph.by_idx.entry(idx).or_insert_with(|| t.clone());
                }
                // Prefer layout font size; fall back to master
                let fs = layout_font_size.or_else(|| master_font_sizes.get(&ph_type).copied());
                if let Some(fs) = fs {
                    lph.by_idx_font_size.entry(idx).or_insert(fs);
                }
                let family = layout_font_family
                    .clone()
                    .or_else(|| master_font_families.get(&ph_type).cloned());
                if let Some(family) = family {
                    lph.by_idx_font_family.entry(idx).or_insert(family);
                }
                // Per-level: layout lstStyle wins per level, else master.
                let level_sizes = merge_level_sizes(
                    &layout_level_sizes,
                    master_level_font_sizes.get(&ph_type).unwrap_or(&[None; 9]),
                );
                if has_any_level_size(&level_sizes) {
                    lph.by_idx_level_sizes.entry(idx).or_insert(level_sizes);
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
                if let Some(ls) = layout_line_spacing {
                    lph.by_idx_line_spacing.entry(idx).or_insert(ls);
                }
                if has_layout_text_inset {
                    lph.by_idx_text_insets
                        .entry(idx)
                        .or_insert(layout_text_insets);
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
            let effective_fs =
                layout_font_size.or_else(|| master_font_sizes.get(&ph_type).copied());
            if let Some(fs) = effective_fs {
                lph.by_type_font_size.entry(ph_type.clone()).or_insert(fs);
            }
            let effective_family = layout_font_family
                .clone()
                .or_else(|| master_font_families.get(&ph_type).cloned());
            if let Some(family) = effective_family {
                lph.by_type_font_family
                    .entry(ph_type.clone())
                    .or_insert(family);
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
            if let Some(v) = layout_space_before {
                lph.by_type_space_before.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(v) = layout_space_after {
                lph.by_type_space_after.entry(ph_type.clone()).or_insert(v);
            }
            if let Some(ls) = layout_line_spacing {
                lph.by_type_line_spacing
                    .entry(ph_type.clone())
                    .or_insert(ls);
            }
            if has_layout_text_inset {
                lph.by_type_text_insets
                    .entry(ph_type.clone())
                    .or_insert(layout_text_insets);
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
    for (ph_type, value) in master_font_sizes {
        lph.by_type_font_size
            .entry(ph_type.clone())
            .or_insert(*value);
    }
    for (ph_type, value) in master_font_families {
        lph.by_type_font_family
            .entry(ph_type.clone())
            .or_insert_with(|| value.clone());
    }
    for (ph_type, value) in master_level_font_sizes {
        lph.by_type_level_sizes
            .entry(ph_type.clone())
            .or_insert(*value);
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
    master_font_sizes: &HashMap<String, f64>,
    master_font_families: &HashMap<String, String>,
    master_level_font_sizes: &HashMap<String, LevelFontSizes>,
    master_level_indents: &HashMap<String, LevelIndents>,
    master_level_bullets: &HashMap<String, LevelBullets>,
    master_anchors: &HashMap<String, String>,
    master_transforms: &HashMap<String, Transform>,
    master_alignments: &HashMap<String, String>,
    master_ea_ln_brk: &HashMap<String, bool>,
    master_space_before: &HashMap<String, i64>,
    master_space_after: &HashMap<String, i64>,
    master_line_spacing: &HashMap<String, f64>,
    theme_source: &(impl PptxThemeSource + ?Sized),
    layout_dir: &str,
    layout_rels: &HashMap<String, String>,
    zip: &mut PptxZip,
) -> ParsedLayout {
    note_layout_master_parse();
    let doc = match parse_preflighted_pptx_xml(layout_xml) {
        Ok(d) => d,
        // Unparseable layout → same as no layout: default placeholders/bg and
        // showMasterSp = true (the slide's own flag still applies downstream).
        Err(_) => return ParsedLayout::default(),
    };
    let root = doc.root_element();

    let placeholders = parse_layout_placeholders(
        root,
        master_font_sizes,
        master_font_families,
        master_level_font_sizes,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_space_before,
        master_space_after,
        master_line_spacing,
        theme_source,
        layout_dir,
        layout_rels,
        zip,
    );

    // Layout-level bg (rels = layout rels, part dir = layout_dir). Verbatim from
    // the former inline layout-bg block in `parse_slide`; the slide decides
    // whether to use it (only when its own bg chain is empty).
    let background: Option<Fill> = child(root, "cSld").and_then(|n| {
        let mut resolve = |rid: &str| -> Option<String> {
            let target = layout_rels.get(rid)?;
            let path = resolve_path(layout_dir, target);
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
    /// pre-baked (logical names → slot hex). Includes font/line/objectDefault
    /// keys exactly as `parse_theme_colors` produced them.
    pub(crate) theme: PptxTheme,
    pub(crate) master_xml: Option<String>,
    pub(crate) master_rels: HashMap<String, String>,
    pub(crate) master_dir: String,
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
    pub(crate) master_font_sizes: HashMap<String, f64>,
    pub(crate) master_font_families: HashMap<String, String>,
    pub(crate) master_level_font_sizes: HashMap<String, LevelFontSizes>,
    pub(crate) master_level_indents: HashMap<String, LevelIndents>,
    pub(crate) master_level_bullets: HashMap<String, LevelBullets>,
    pub(crate) master_anchors: HashMap<String, String>,
    pub(crate) master_transforms: HashMap<String, Transform>,
    pub(crate) master_alignments: HashMap<String, String>,
    pub(crate) master_ea_ln_brk: HashMap<String, bool>,
    pub(crate) master_space_before: HashMap<String, i64>,
    pub(crate) master_space_after: HashMap<String, i64>,
    pub(crate) master_line_spacing: HashMap<String, f64>,
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
    /// Master per-level bullet colors re-resolved against `theme`.
    pub(crate) master_level_bullets: HashMap<String, LevelBullets>,
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
/// TODO: themeOverride (slide/layout `/themeOverride`, ECMA-376 §14.2.7) is not
/// yet honored — overrides on the layout or slide would replace parts of the
/// master theme. Out of scope for per-slide master resolution.
pub(crate) fn build_master_bundle(
    master_path: &str,
    fallback_theme: &PptxTheme,
    zip: &mut PptxZip,
) -> ParsedMaster {
    let master_xml_opt: Option<String> = if master_path.is_empty() {
        None
    } else {
        read_zip_str(zip, master_path).ok()
    };

    let master_dir: String = master_path
        .rsplit_once('/')
        .map(|(dir, _)| dir.to_owned())
        .unwrap_or_else(|| "ppt/slideMasters".to_owned());

    // Master rels: `<master_dir>/_rels/<file>.rels`.
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
        find_rel_target_by_type(&master_rels_xml, "/theme").map(|t| resolve_path(&master_dir, &t));
    let mut theme = theme_path
        .as_deref()
        .map(|path| parse_theme_part(path, zip))
        .unwrap_or_else(|| fallback_theme.clone());
    // Bake the master's <p:clrMap> logical-name → slot mapping into the theme.
    bake_clr_map(&mut theme, master_xml_opt.as_deref());

    let master_smartart_drawings: HashMap<String, String> =
        build_smartart_drawings(&master_rels_xml, &master_dir, zip);

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
            let path = resolve_path(&master_dir, target);
            // Existence check only — central-directory lookup, no inflate
            // (former `read_zip_bytes` decompressed the entry just to discard it).
            zip.index_for_name(&path)?;
            Some(path)
        };
        parse_background(c_sld, &theme, &mut resolve)
    });

    let master_font_sizes = master_root.map(parse_master_font_sizes).unwrap_or_default();
    let master_font_families = master_root
        .map(|root| parse_master_font_families(root, &theme))
        .unwrap_or_default();
    let master_level_font_sizes = master_root
        .map(parse_master_level_font_sizes)
        .unwrap_or_default();
    let master_level_indents = master_root
        .map(parse_master_level_indents)
        .unwrap_or_default();
    let master_level_bullets = master_root
        .map(|root| parse_master_level_bullets(root, &theme, &master_rels, &master_dir, zip))
        .unwrap_or_default();
    let master_anchors = master_root.map(parse_master_anchors).unwrap_or_default();
    let master_transforms = master_root.map(parse_master_transforms).unwrap_or_default();
    let master_alignments = master_root.map(parse_master_alignments).unwrap_or_default();
    let master_ea_ln_brk = master_root.map(parse_master_ea_ln_brk).unwrap_or_default();
    let (master_space_before, master_space_after, master_line_spacing) = master_root
        .map(parse_master_txstyle_spacing)
        .unwrap_or_default();
    let (master_bold, master_italic, master_caps, master_reflection) = master_root
        .map(parse_master_txstyle_run_properties)
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
            &master_dir,
            &master_rels,
            &master_smartart_drawings,
            &theme,
            zip,
            &mut master_decorative,
        );
    }

    ParsedMaster {
        theme,
        master_xml: master_xml_opt,
        master_rels,
        master_dir,
        master_smartart_drawings,
        master_bg,
        master_decorative,
        master_font_sizes,
        master_font_families,
        master_level_font_sizes,
        master_level_indents,
        master_level_bullets,
        master_anchors,
        master_transforms,
        master_alignments,
        master_ea_ln_brk,
        master_space_before,
        master_space_after,
        master_line_spacing,
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
    use crate::text::{BuMarker, BulletProps};
    use std::io::Cursor;

    fn empty_zip() -> PptxZip {
        let writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let cursor = writer.finish().unwrap();
        PptxZip::new(cursor).unwrap()
    }

    fn parse_layout_with_master(
        layout_shape: &str,
        master_font_sizes: &HashMap<String, f64>,
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
            master_font_sizes,
            &HashMap::<String, String>::new(),
            &HashMap::<String, LevelFontSizes>::new(),
            &HashMap::<String, LevelIndents>::new(),
            &HashMap::<String, LevelBullets>::new(),
            &HashMap::<String, String>::new(),
            &HashMap::<String, Transform>::new(),
            &HashMap::<String, String>::new(),
            &HashMap::<String, bool>::new(),
            &HashMap::<String, i64>::new(),
            &HashMap::<String, i64>::new(),
            &HashMap::<String, f64>::new(),
            &HashMap::new(),
            "ppt/slideLayouts",
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
            "ppt/slides",
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
    #[test]
    fn master_without_tx_styles_uses_powerpoint_placeholder_font_defaults() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();

        let sizes = parse_master_font_sizes(doc.root_element());
        assert_eq!(sizes.get("title"), Some(&44.0));
        assert_eq!(sizes.get("ctrTitle"), Some(&44.0));
        assert_eq!(sizes.get("body"), Some(&28.0));
        assert_eq!(sizes.get("subTitle"), Some(&28.0));
        assert_eq!(sizes.get("obj"), Some(&18.0));
        assert_eq!(sizes.get(""), Some(&18.0));
        assert_eq!(sizes.get("dt"), None);
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
        let master_sizes = HashMap::from([("body".to_owned(), 28.0), ("obj".to_owned(), 18.0)]);
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
        assert_eq!(shape.text_body.unwrap().default_font_size, Some(18.0));
    }

    #[test]
    fn missing_tx_styles_supply_only_observed_level_one_body_list_defaults() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let indents = parse_master_level_indents(doc.root_element());
        let mut zip = empty_zip();
        let bullets = parse_master_level_bullets(
            doc.root_element(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideMasters",
            &mut zip,
        );

        assert_eq!(indents["body"][0].mar_l, Some(228_600));
        assert_eq!(indents["body"][0].indent, Some(-228_600));
        assert_eq!(indents["subTitle"][0].mar_l, Some(228_600));
        assert!(!indents.contains_key("obj"));
        match bullets["body"][0].resolve() {
            Bullet::Char { ch, .. } => assert_eq!(ch, "•"),
            other => panic!("expected implicit body bullet, got {other:?}"),
        }
        match bullets["subTitle"][0].resolve() {
            Bullet::Char { ch, .. } => assert_eq!(ch, "•"),
            other => panic!("expected implicit subtitle bullet, got {other:?}"),
        }
        assert!(!bullets.contains_key("obj"));
    }

    #[test]
    fn authored_master_tx_styles_override_powerpoint_placeholder_defaults() {
        let xml = r#"<p:sldMaster
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <p:cSld><p:spTree/></p:cSld>
          <p:txStyles>
            <p:titleStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></p:titleStyle>
            <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle>
            <p:otherStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:otherStyle>
          </p:txStyles>
        </p:sldMaster>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();

        let sizes = parse_master_font_sizes(doc.root_element());

        assert_eq!(sizes.get("title"), Some(&36.0));
        assert_eq!(sizes.get("body"), Some(&24.0));
        assert_eq!(sizes.get("dt"), Some(&12.0));
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
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &master_bullets,
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            &HashMap::new(),
            "ppt/slideLayouts",
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
              <a:bodyPr lIns="216000" tIns="72000" rIns="216000" bIns="72000" anchor="ctr"/>
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

        let local_left_override = slide.replace("<a:bodyPr/>", "<a:bodyPr lIns=\"0\"/>");
        let shape = parse_slide_shape(&local_left_override, &placeholders);
        let body = shape.text_body.expect("placeholder text body");
        assert_eq!(body.l_ins, 0);
        assert_eq!(body.t_ins, 72_000);
        assert_eq!(body.r_ins, 216_000);
        assert_eq!(body.b_ins, 72_000);
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
