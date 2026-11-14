//! Shared DrawingML blip helpers used by the docx, pptx and xlsx parsers.
//!
//! Every OOXML host references embedded media through the same DrawingML
//! `<a:blip>` element (ECMA-376 §20.1.8.13). A blip's raster image rides in its
//! `@r:embed` relationship; when the picture is a vector graphic, Microsoft's
//! 2016 SVG extension (MS-ODRAWXML) nests an `<asvg:svgBlip r:embed="…">` inside
//! the blip's `<a:extLst>`, and the raster `@r:embed` becomes a *fallback* for
//! SVG-incapable clients. These helpers were previously re-implemented (or, for
//! the SVG extension, only implemented in pptx) per parser; sharing them keeps
//! the three formats' blip handling identical.

use crate::ns::relationships;
use roxmltree::Node;
use serde::{Deserialize, Serialize};

/// Transitional relationships namespace (`r:`) — where a blip's `embed` / `link`
/// rIds live. Runtime lookups accept the Strict URI too (see [`blip_embed_rid`]);
/// this constant is kept for the crate's test fixtures.
pub const R_NS: &str = relationships::TRANSITIONAL;

/// Microsoft 2016 SVG extension URI (MS-ODRAWXML). An `<a:blip>` wrapping a
/// vector image carries `<a:extLst><a:ext uri="{96DAC541-…}"><asvg:svgBlip
/// r:embed="…"/></a:ext></a:extLst>`.
pub const SVG_BLIP_EXT_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
pub const SVG_BLIP_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

/// Resolve a blip-like node's `r:embed` relationship id (the raster image, or
/// an `svgBlip`'s vector target). Reads the `embed` attribute in the
/// relationships namespace — Transitional or Strict (ISO/IEC 29500) — tolerating
/// the literal `r:embed` form some producers emit without binding the namespace.
pub fn blip_embed_rid(node: &Node<'_, '_>) -> Option<String> {
    node.attribute((relationships::TRANSITIONAL, "embed"))
        .or_else(|| node.attribute((relationships::STRICT, "embed")))
        .or_else(|| node.attribute("r:embed"))
        .map(str::to_string)
}

/// Resolve the relationship id of the vector original from an `<a:blip>`'s
/// `asvg:svgBlip` extension. Both the registered extension URI and Microsoft
/// SVG namespace are required; an unrelated extension may legally reuse the
/// same local name. The producer's prefix (`asvg:` etc.) does not matter. Returns
/// `None` when the blip carries no svgBlip extension (the common, raster-only
/// case).
pub fn svg_blip_rid(blip: Node<'_, '_>) -> Option<String> {
    blip.children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")?
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "ext")
        .filter(|ext| ext.attribute("uri") == Some(SVG_BLIP_EXT_URI))
        .find_map(|ext| {
            ext.children().find(|n| {
                n.is_element()
                    && n.tag_name().name() == "svgBlip"
                    && n.tag_name().namespace() == Some(SVG_BLIP_NS)
            })
        })
        .and_then(|svg_blip| blip_embed_rid(&svg_blip))
}

/// Map a part name / path extension to its MIME type, covering the image, audio
/// and video parts OOXML blips reference. Unknown extensions fall back to
/// `application/octet-stream`. This is the single source of truth for all three
/// parsers — notably it is the only place `svg` ⇒ `image/svg+xml` is decided.
///
/// The extension is read from the part name's equivalence key
/// ([`crate::rels::part_name_equivalence_key`], ECMA-376 Part 2 §6.2.2.3 with
/// RFC 3986 §6.2.2), so a stored item name such as `image1.%73vg` or
/// `IMAGE1.SVG` infers the same type as `image1.svg`.
pub fn mime_from_ext(path: &str) -> &'static str {
    let key = crate::rels::part_name_equivalence_key(path);
    match key.rsplit('.').next().unwrap_or("") {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "tif" | "tiff" => "image/tiff",
        // Windows Metafile / Enhanced Metafile. Office embeds these for charts
        // and diagrams; the renderer rasterizes WMF via a minimal player and
        // skips EMF (a follow-up). The conventional MIME types per IANA / Windows.
        "wmf" => "image/wmf",
        "emf" => "image/emf",
        "mp3" => "audio/mpeg",
        "m4a" => "audio/mp4",
        "wav" => "audio/wav",
        "aac" => "audio/aac",
        "wma" => "audio/x-ms-wma",
        "flac" => "audio/flac",
        "ogg" | "oga" => "audio/ogg",
        "mp4" | "m4v" => "video/mp4",
        "mov" | "qt" => "video/quicktime",
        "avi" => "video/x-msvideo",
        "wmv" => "video/x-ms-wmv",
        "mpg" | "mpeg" => "video/mpeg",
        "3gp" => "video/3gpp",
        "mkv" => "video/x-matroska",
        "webm" => "video/webm",
        "ogv" => "video/ogg",
        _ => "application/octet-stream",
    }
}

/// ECMA-376 §20.1.8.55 `<a:srcRect>` source-image crop (DrawingML
/// `CT_RelativeRect`). Each edge inset is a signed fraction of the *source*
/// bitmap, measured from that edge, so the visible source region is
/// `[l, t, 1−r, 1−b]`. The raw attributes are `ST_Percentage` in 1000ths of a
/// percent; the parser divides by 100000 to a fraction so renderers need no unit
/// knowledge. `ST_Percentage` has no fixed range; absent edges default to `0`.
/// Shared by the docx, pptx and xlsx parsers so the three formats crop identically.
#[derive(Serialize, Deserialize, Debug, Clone, Copy, Default, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(rename_all = "camelCase")]
pub struct SrcRect {
    pub l: f64,
    pub t: f64,
    pub r: f64,
    pub b: f64,
}

/// Parse `<a:srcRect l t r b>` from a `<*:blipFill>` node (the parent of the
/// `<a:blip>`), matched by namespace-local name so a `p:`, `a:` or `xdr:`
/// blipFill all work. Each edge is divided by 100000 to a fraction; an absent
/// edge defaults to `0`. Returns `None` when there is no `srcRect` or all four
/// edges are zero (no crop), so an uncropped picture never forces a renderer
/// onto the sub-rectangle draw path.
pub fn parse_src_rect(blip_fill: Node<'_, '_>) -> Option<SrcRect> {
    let sr = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "srcRect")?;
    let read = |name: &str| -> f64 {
        sr.attribute(name)
            .and_then(crate::units::drawingml_percentage_to_fraction)
            .unwrap_or(0.0)
    };
    let rect = SrcRect {
        l: read("l"),
        t: read("t"),
        r: read("r"),
        b: read("b"),
    };
    if rect.l.abs() < 1e-9 && rect.t.abs() < 1e-9 && rect.r.abs() < 1e-9 && rect.b.abs() < 1e-9 {
        None
    } else {
        Some(rect)
    }
}

/// Parse `<a:blip><a:alphaModFix amt="…"/></a:blip>` from a `<*:blipFill>` node
/// (ECMA-376 §20.1.8.6 CT_AlphaModulateFixedEffect). `amt` is an
/// ST_PositivePercentage in 1000ths of a percent, so `amt/100000` is the alpha
/// scale fraction (0.0–1.0). The effect multiplies the picture's opacity by that
/// fraction — the renderer applies it via `globalAlpha`.
///
/// Returns `None` when there is no `alphaModFix` OR when the amount is ≥ 100%
/// (fully opaque = nothing to apply), so an unaffected picture never forces the
/// renderer onto the alpha path. Negative amounts clamp to 0. Shared by the
/// pptx and xlsx parsers (docx can carry it too) so the three formats read the
/// blip alpha identically.
pub fn parse_blip_alpha(blip_fill: Node<'_, '_>) -> Option<f64> {
    let blip = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "blip")?;
    let mut found = false;
    let mut frac = 1.0;
    for effect in blip
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "alphaModFix")
    {
        let amt = effect
            .attribute("amt")
            .map(crate::units::drawingml_percentage_to_fraction)
            .unwrap_or(Some(1.0))?;
        frac *= amt.max(0.0);
        found = true;
    }
    if !found {
        return None;
    }
    if frac >= 1.0 {
        None
    } else {
        Some(frac)
    }
}

/// A `<a:duotone>` image effect (ECMA-376 §20.1.8.23) resolved to its two
/// endpoint colours. `clr1` is the first `EG_ColorChoice` child (the dark
/// endpoint, luminance 0), `clr2` the second (the light endpoint, luminance 1) —
/// the spec says the effect "combines clr1 and clr2 through a linear
/// interpolation" per pixel, so pixel luminance is the interpolation factor.
/// Both are 6- or 8-char uppercase hex WITHOUT a `#` (RRGGBB or RRGGBBAA);
/// any per-colour transforms
/// (lumMod/lumOff/tint/satMod/shade/…) are already baked in by
/// [`crate::color::parse_color_node`]. Shared across the docx/pptx/xlsx parsers.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(rename_all = "camelCase")]
pub struct Duotone {
    /// First colour child — dark endpoint. 6-char hex, no `#`.
    pub clr1: String,
    /// Second colour child — light endpoint. 6-char hex, no `#`.
    pub clr2: String,
}

/// Parse `<a:blipFill>`'s `<a:duotone>` (§20.1.8.23), resolving its two
/// `EG_ColorChoice` children through the shared DrawingML colour grammar
/// ([`crate::color::parse_color_node`]) so all per-colour transforms are applied.
/// `resolver` supplies the host's `<a:schemeClr>` palette lookup and `tint_mode`
/// selects the Word/PowerPoint `<a:tint>` interpretation, exactly as every other
/// colour path in the parser does.
///
/// `<a:duotone>` is a CT_Blip child. For legacy callers this helper also
/// tolerates a direct blipFill child (outside the schema). Returns `None` when there is no
/// duotone or fewer than two resolvable colours (a malformed effect is dropped
/// rather than half-applied). The returned hexes carry NO `#` and are uppercase.
pub fn parse_blip_duotone(
    blip_fill: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
) -> Option<Duotone> {
    // `<a:duotone>` lives on the blip itself (a blip-effect child), with the
    // srcRect/stretch/tile as its siblings. Some producers place it directly
    // under blipFill; accept either by searching descendants of blipFill's
    // immediate children shallowly.
    let duotone = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "duotone")
        .or_else(|| {
            blip_fill
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "blip")
                .and_then(|blip| {
                    blip.children()
                        .find(|n| n.is_element() && n.tag_name().name() == "duotone")
                })
        })?;
    // The two EG_ColorChoice children, in document order. Each child IS a colour
    // element (srgbClr/schemeClr/prstClr/…) — not a container wrapping one — so it
    // is classified with `color_source_from_element` and resolved (transforms and
    // all) via `resolve_color_source`, reusing the exact colour grammar every
    // other path uses.
    let mut colors = duotone
        .children()
        .filter(|n| n.is_element())
        .filter_map(|node| crate::color::color_source_from_element(node))
        .filter_map(|source| crate::color::resolve_color_source(source, resolver, tint_mode));
    let clr1 = colors.next()?;
    let clr2 = colors.next()?;
    Some(Duotone {
        clr1: clr1.to_uppercase(),
        clr2: clr2.to_uppercase(),
    })
}

/// Ordered CT_Blip effects (ECMA-376 §20.1.8.13). Each entry owns all of
/// its parameters; malformed effects remain explicit diagnostics.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BlipEffect {
    /// §20.1.8.34 `<a:grayscl>`.
    Grayscale,
    /// §20.1.8.11 `<a:biLevel thresh>`, threshold as a 0–1 fraction.
    #[serde(rename_all = "camelCase")]
    BiLevel {
        thresh: f64,
    },
    /// §20.1.8.16 `<a:clrChange>`: colours as 6-char uppercase hex (no `#`)
    /// with their `<a:alpha>` as 0–1 fractions; `use_alpha` is `@useA`.
    #[serde(rename_all = "camelCase")]
    ColorChange {
        from: String,
        from_alpha: f64,
        to: String,
        to_alpha: f64,
        use_alpha: bool,
    },
    /// §20.1.8.42 `<a:lum bright contrast>` as signed fractions (100000 =
    /// 1.0). The rendering formula comes from PowerPoint's output (see the
    /// renderer).
    Luminance {
        bright: f64,
        contrast: f64,
    },
    /// This duotone's own endpoints, not a reference to a global palette.
    #[serde(rename_all = "camelCase")]
    Duotone {
        clr1: String,
        clr2: String,
    },
    /// §20.1.8.1–4, 6, 8: alpha effects operate at their authored position.
    AlphaBiLevel {
        thresh: f64,
    },
    AlphaCeiling,
    AlphaFloor,
    AlphaInv {
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
    },
    /// §20.1.8.5 + §20.1.10.22: recursively evaluated alpha mask.
    AlphaMod {
        container: EffectContainer,
    },
    /// §20.1.8.15: physical radius in EMUs; grow defaults true.
    Blur {
        radius: f64,
        grow: bool,
    },
    /// §20.1.8.39: hue degrees and relative S/L fractions.
    Hsl {
        hue: f64,
        sat: f64,
        lum: f64,
    },
    /// §20.1.8.60: hue degrees and signed color shift.
    Tint {
        hue: f64,
        amt: f64,
    },
    /// §20.1.8.29 + §20.1.10.11: one shared DrawingML fill recipe.
    FillOverlay {
        fill: EffectFill,
        blend: String,
    },
    /// §20.1.8.20 / 24, valid within a recursive effect container.
    Container {
        container: EffectContainer,
    },
    Reference {
        reference: String,
    },
    /// Part-local DAG recipes reached by name. Metadata only; `work` records
    /// bounded XML lookup visits and is charged to the existing effect budget.
    Definitions {
        containers: Vec<EffectContainer>,
        work: u64,
    },
    /// Parser-side recursion crossed the same evaluated-step resource policy.
    EffectCountLimit {
        observed: usize,
    },
    // EG_Effect members legal inside alphaMod's container (§A.4.1).
    AlphaOutset {
        radius: f64,
    },
    Blend {
        container: EffectContainer,
        blend: String,
    },
    Fill {
        fill: EffectFill,
    },
    Glow {
        glow: crate::effect::Glow,
    },
    Shadow {
        shadow: crate::effect::Shadow,
        inner: bool,
    },
    SoftEdge {
        radius: f64,
    },
    Reflection {
        reflection: BlipReflection,
    },
    Transform {
        sx: f64,
        sy: f64,
        kx: f64,
        ky: f64,
        tx: f64,
        ty: f64,
    },
    RelativeOffset {
        tx: f64,
        ty: f64,
    },
    PresetShadow {
        shadows: Vec<crate::effect::Shadow>,
    },
    AlphaModFix {
        amt: f64,
    },
    AlphaRepl {
        alpha: f64,
    },
    /// §20.1.8.18 replaces RGB, preserving the source alpha.
    ColorReplace {
        color: String,
    },
    /// Explicit fail-closed diagnostic, consumed before decode/paint. The
    /// unsupported entry is retained even when it is the only blip effect.
    Unsupported {
        effect: String,
    },
}

/// CT_ReflectionEffect's complete geometry, extending the shared shape
/// descriptor with the attributes a recursive mask must not discard.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(rename_all = "camelCase")]
pub struct BlipReflection {
    #[serde(flatten)]
    pub base: crate::effect::Reflection,
    pub kx: f64,
    pub ky: f64,
    pub algn: String,
    pub fade_dir: f64,
    pub rot_with_shape: bool,
}

/// §20.1.8.20, §20.1.10.22. Siblings each consume the parent;
/// trees consume the previous result. Names remain owned graph facts.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(rename_all = "camelCase")]
pub struct EffectContainer {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub effects: Vec<BlipEffect>,
}

/// Fill recipes match core's shared Fill model. Geometry and palette parsing
/// use fill.rs; there is no second gradient/pattern implementation here.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
#[cfg_attr(feature = "schema", derive(schemars::JsonSchema))]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub enum EffectFill {
    None,
    Solid {
        color: String,
    },
    #[serde(rename_all = "camelCase")]
    Gradient {
        stops: Vec<crate::fill::GradStop>,
        angle: f64,
        grad_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_to_rect: Option<crate::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        tile_rect: Option<crate::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
    },
    Pattern {
        fg: String,
        bg: String,
        preset: String,
    },
    #[serde(rename_all = "camelCase")]
    Image {
        image_path: String,
        mime_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        svg_image_path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        src_rect: Option<SrcRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_rect: Option<crate::fill::FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        tile: Option<crate::fill::TileInfo>,
        #[serde(skip_serializing_if = "Option::is_none")]
        dpi: Option<f64>,
        stretch: bool,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
        blip_effects: Vec<BlipEffect>,
    },
}

// Resource policy shared with core/image/pixel-budget.ts, not schema limits.
const MAX_IMAGE_EFFECT_STEPS: usize = 16;
const MAX_IMAGE_EFFECT_LOOKUP_WORK: u64 = 1 << 27;
#[derive(Default)]
struct EffectParseContext {
    active: std::collections::HashSet<roxmltree::NodeId>,
    steps: usize,
    work: u64,
    lookup_work: u64,
    failure: Option<BlipEffect>,
}

impl EffectParseContext {
    // One ledger covers all XML-to-recipe edges: inherited grpFill, nested
    // blipFill, inline alphaMod/blend/cont and named effectDag. Checking only
    // named definitions leaves inheritance cycles and deep inline XML unsafe.
    // These are library resource ceilings, not ECMA-376 schema restrictions.
    fn enter(&mut self, node: Node<'_, '_>) -> bool {
        if self.failure.is_some() {
            return false;
        }
        if self.active.contains(&node.id()) {
            self.failure = Some(BlipEffect::Unsupported {
                effect: "cyclic effect reference or inherited fill".into(),
            });
            return false;
        }
        if self.active.len() >= MAX_IMAGE_EFFECT_STEPS {
            self.failure = Some(BlipEffect::EffectCountLimit {
                observed: MAX_IMAGE_EFFECT_STEPS + 1,
            });
            return false;
        }
        self.active.insert(node.id());
        true
    }
    fn charge(&mut self, units: u64) -> bool {
        if self.failure.is_some() {
            return false;
        }
        self.work = self.work.saturating_add(units);
        if self.work.saturating_add(self.lookup_work) > MAX_IMAGE_EFFECT_LOOKUP_WORK
            || crate::resource::charge_image_effect_work(units, MAX_IMAGE_EFFECT_LOOKUP_WORK)
                .is_err()
        {
            self.failure = Some(BlipEffect::Definitions {
                containers: Vec::new(),
                work: MAX_IMAGE_EFFECT_LOOKUP_WORK + 1,
            });
            return false;
        }
        true
    }
}

/// Parse CT_Blip effects without losing order or silently dropping effects.
/// A single duotone retains the legacy fast path; every other chain is complete.
pub fn parse_blip_effects(
    blip_fill: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
) -> Vec<BlipEffect> {
    let Some(blip) = blip_fill
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "blip")
    else {
        return Vec::new();
    };
    let context = &mut EffectParseContext::default();
    let mut effects = parse_effect_children(blip, resolver, tint_mode, context);
    // CT_Blip's immediate effects form a sequential chain. Only here can
    // identities be omitted without deleting a separately composited branch.
    // §20.1.8.6 / §20.1.10.22: a sibling identity contributes parent coverage.
    // Keep all recursive containers (including named recipes) structurally intact.
    effects.retain(|effect| {
        !matches!(effect,
        BlipEffect::AlphaModFix { amt } if *amt == 1.0)
            && !matches!(effect, BlipEffect::Luminance { bright, contrast }
            if *bright == 0.0 && *contrast == 0.0)
    });
    complete_named_effects(&mut effects, blip, resolver, tint_mode, context);
    if let Some(failure) = context.failure.take() {
        // Preserve a transportable failure even through hosts whose parser API
        // stores effects instead of returning Result. Core preflight turns this
        // into unsupported-image-effect or the typed image-effect quota error.
        return vec![failure];
    }
    if context.work > 0 {
        effects.push(BlipEffect::Definitions {
            containers: Vec::new(),
            work: context.work,
        });
    }
    // Only outer hosts also parse a legacy standalone duotone palette.
    // Recursive fills/containers have no separate palette and keep the entry.
    if effects.len() == 1
        && matches!(effects[0], BlipEffect::Duotone { ref clr1, ref clr2 } if clr1.len() == 6 && clr2.len() == 6)
    {
        Vec::new()
    } else {
        effects
    }
}

fn parse_effect_children(
    blip: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
    context: &mut EffectParseContext,
) -> Vec<BlipEffect> {
    if !context.enter(blip) {
        return Vec::new();
    }
    let mut effects = Vec::new();

    for node in blip.children().filter(|n| n.is_element()) {
        if context.failure.is_some() {
            break;
        }
        if node.tag_name().name() == "extLst" {
            continue;
        }
        context.steps += 1;
        if context.steps > MAX_IMAGE_EFFECT_STEPS {
            context.failure = Some(BlipEffect::EffectCountLimit {
                observed: context.steps,
            });
            break;
        }
        let effect = parse_effect_node(node, resolver, tint_mode, context);
        effects.push(effect.unwrap_or_else(|| BlipEffect::Unsupported {
            effect: node.tag_name().name().to_owned(),
        }));
    }
    context.active.remove(&blip.id());
    effects
}

// Recursive dispatch has a small stack frame. Keeping nonrecursive leaf
// parsing separate avoids multiplying its large enum-construction workspace
// by the allowed nesting depth (including native debug and WASM builds).
fn parse_effect_node(
    node: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
    context: &mut EffectParseContext,
) -> Option<BlipEffect> {
    match node.tag_name().name() {
        "alphaMod" => node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "cont")
            .map(|cont| BlipEffect::AlphaMod {
                container: parse_container(cont, resolver, tint_mode, context),
            }),
        "cont" | "effectDag" => Some(BlipEffect::Container {
            container: parse_container(node, resolver, tint_mode, context),
        }),
        "fillOverlay" => node.attribute("blend").and_then(|blend| {
            parse_effect_fill(node, resolver, tint_mode, context).map(|fill| {
                BlipEffect::FillOverlay {
                    fill,
                    blend: blend.into(),
                }
            })
        }),
        "blend" => node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "cont")
            .zip(node.attribute("blend"))
            .map(|(cont, blend)| BlipEffect::Blend {
                container: parse_container(cont, resolver, tint_mode, context),
                blend: blend.into(),
            }),
        "fill" => parse_effect_fill(node, resolver, tint_mode, context)
            .map(|fill| BlipEffect::Fill { fill }),
        _ => parse_leaf_effect(node, resolver, tint_mode),
    }
}

fn parse_leaf_effect(
    node: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
) -> Option<BlipEffect> {
    let color = |container: Option<Node<'_, '_>>| -> Option<(String, f64)> {
        let element = container?.children().find(|n| n.is_element())?;
        let source = crate::color::color_source_from_element(element)?;
        // The shared resolver appends the resolved alpha transforms as a
        // trailing byte (RRGGBBAA) when the colour is not opaque.
        let hex = crate::color::resolve_color_source(source, resolver, tint_mode)?.to_uppercase();
        let alpha = match hex.len() {
            6 => 1.0,
            8 => f64::from(u8::from_str_radix(&hex[6..8], 16).ok()?) / 255.0,
            _ => return None,
        };
        Some((hex[..6].to_owned(), alpha))
    };
    // Defaults apply only to absent attributes (§A.4.1), never malformed
    // authored values. The shared shape-effect helpers are intentionally
    // tolerant; recursive CT_Blip must preserve a typed failure instead.
    for attribute in node.attributes() {
        let key = attribute.name();
        let valid = match key {
            "rad" | "blurRad" | "dist" | "tx" | "ty"
                if !matches!(node.tag_name().name(), "relOff") =>
            {
                crate::units::coordinate_to_emu(attribute.value()).is_some()
            }
            "amt" | "thresh" | "a" | "sat" | "lum" | "bright" | "contrast" | "sx" | "sy"
            | "stA" | "stPos" | "endA" | "endPos" | "tx" | "ty" => {
                crate::units::drawingml_percentage_to_fraction(attribute.value())
                    .is_some_and(f64::is_finite)
            }
            "hue" | "dir" | "fadeDir" | "kx" | "ky" => {
                attribute.value().parse::<f64>().is_ok_and(f64::is_finite)
            }
            "grow" | "useA" | "rotWithShape" => {
                matches!(attribute.value(), "0" | "1" | "true" | "false")
            }
            _ => true,
        };
        if !valid {
            return None;
        }
    }
    if matches!(
        node.tag_name().name(),
        "glow" | "innerShdw" | "outerShdw" | "prstShdw"
    ) && color(Some(node)).is_none()
    {
        return None;
    }
    if node.tag_name().name() == "softEdge" && node.attribute("rad").is_none() {
        return None;
    }
    match node.tag_name().name() {
        "grayscl" => Some(BlipEffect::Grayscale),
        "biLevel" => node
            .attribute("thresh")
            .and_then(crate::units::drawingml_percentage_to_fraction)
            .map(|thresh| BlipEffect::BiLevel {
                thresh: thresh.clamp(0.0, 1.0),
            }),
        "clrChange" => {
            let child = |name: &str| {
                node.children()
                    .find(|n| n.is_element() && n.tag_name().name() == name)
            };
            match (color(child("clrFrom")), color(child("clrTo"))) {
                (Some((from, from_alpha)), Some((to, to_alpha))) => Some(BlipEffect::ColorChange {
                    from,
                    from_alpha,
                    to,
                    to_alpha,
                    use_alpha: node
                        .attribute("useA")
                        .is_some_and(|v| v == "1" || v == "true"),
                }),
                _ => None,
            }
        }
        // §20.1.8.42 CT_LuminanceEffect: ST_FixedPercentage bright and
        // contrast, both defaulting to 0. Retain identity sibling branches.
        "lum" => {
            let fraction = |name: &str| match node.attribute(name) {
                None => Some(0.0),
                Some(value) => crate::units::drawingml_percentage_to_fraction(value)
                    .map(|value| value.clamp(-1.0, 1.0)),
            };
            match (fraction("bright"), fraction("contrast")) {
                (Some(bright), Some(contrast)) => Some(BlipEffect::Luminance { bright, contrast }),
                _ => None,
            }
        }
        "duotone" => {
            let mut colors = node
                .children()
                .filter(|n| n.is_element())
                .filter_map(crate::color::color_source_from_element)
                .filter_map(|source| {
                    crate::color::resolve_color_source(source, resolver, tint_mode)
                });
            match (colors.next(), colors.next()) {
                (Some(clr1), Some(clr2)) => Some(BlipEffect::Duotone {
                    clr1: clr1.to_uppercase(),
                    clr2: clr2.to_uppercase(),
                }),
                _ => None,
            }
        }
        "alphaBiLevel" => node
            .attribute("thresh")
            .and_then(crate::units::drawingml_percentage_to_fraction)
            .map(|thresh| BlipEffect::AlphaBiLevel {
                thresh: thresh.clamp(0.0, 1.0),
            }),
        "alphaCeiling" => Some(BlipEffect::AlphaCeiling),
        "alphaFloor" => Some(BlipEffect::AlphaFloor),
        "alphaInv" => {
            if node.children().any(|n| n.is_element()) {
                color(Some(node)).map(|(color, _)| BlipEffect::AlphaInv { color: Some(color) })
            } else {
                Some(BlipEffect::AlphaInv { color: None })
            }
        }
        "effect" => node.attribute("ref").map(|r| BlipEffect::Reference {
            reference: r.into(),
        }),
        "blur" => {
            let radius = node
                .attribute("rad")
                .map(|v| crate::units::coordinate_to_emu(v).map(|n| n as f64))
                .unwrap_or(Some(0.0));
            radius.map(|radius| BlipEffect::Blur {
                radius: radius.max(0.0),
                grow: !matches!(node.attribute("grow"), Some("0" | "false")),
            })
        }
        "hsl" => {
            let read = |name| {
                node.attribute(name)
                    .map(crate::units::drawingml_percentage_to_fraction)
                    .unwrap_or(Some(0.0))
            };
            let hue = node
                .attribute("hue")
                .map(|v| v.parse::<f64>().ok().map(|n| n / 60000.0))
                .unwrap_or(Some(0.0));
            match (hue, read("sat"), read("lum")) {
                (Some(hue), Some(sat), Some(lum)) => Some(BlipEffect::Hsl { hue, sat, lum }),
                _ => None,
            }
        }
        "tint" => {
            let hue = node
                .attribute("hue")
                .map(|v| v.parse::<f64>().ok().map(|n| n / 60000.0))
                .unwrap_or(Some(0.0));
            let amt = node
                .attribute("amt")
                .map(crate::units::drawingml_percentage_to_fraction)
                .unwrap_or(Some(0.0));
            match (hue, amt) {
                (Some(hue), Some(amt)) => Some(BlipEffect::Tint { hue, amt }),
                _ => None,
            }
        }
        "alphaOutset" => Some(BlipEffect::AlphaOutset {
            radius: node
                .attribute("rad")
                .and_then(crate::units::coordinate_to_emu)
                .unwrap_or(0) as f64,
        }),
        "glow" => Some(BlipEffect::Glow {
            glow: crate::effect::parse_glow_effect(node, resolver, tint_mode),
        }),
        "innerShdw" | "outerShdw" => Some(BlipEffect::Shadow {
            shadow: crate::effect::parse_shadow_effect(node, resolver, tint_mode),
            inner: node.tag_name().name() == "innerShdw",
        }),
        "softEdge" => Some(BlipEffect::SoftEdge {
            radius: crate::effect::parse_soft_edge_effect(node).radius as f64,
        }),
        "reflection" => Some(BlipEffect::Reflection {
            reflection: BlipReflection {
                base: {
                    let mut reflection = crate::effect::parse_reflection_effect(node);
                    // CT_ReflectionEffect@sy defaults to +100% (§A.4.1).
                    // Keep the established shape painter's omitted-mirror
                    // policy outside the recursive CT_Blip evaluator.
                    reflection.sy = node
                        .attribute("sy")
                        .and_then(crate::units::drawingml_percentage_to_fraction)
                        .unwrap_or(1.0);
                    reflection
                },
                kx: node
                    .attribute("kx")
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(0.0)
                    / 60000.0,
                ky: node
                    .attribute("ky")
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(0.0)
                    / 60000.0,
                algn: node.attribute("algn").unwrap_or("b").into(),
                fade_dir: node
                    .attribute("fadeDir")
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(5400000.0)
                    / 60000.0,
                rot_with_shape: !matches!(node.attribute("rotWithShape"), Some("false" | "0")),
            },
        }),
        "prstShdw" => parse_preset_shadow(node, resolver, tint_mode)
            .map(|shadows| BlipEffect::PresetShadow { shadows }),
        "relOff" => Some(BlipEffect::RelativeOffset {
            tx: node
                .attribute("tx")
                .and_then(crate::units::drawingml_percentage_to_fraction)
                .unwrap_or(0.0),
            ty: node
                .attribute("ty")
                .and_then(crate::units::drawingml_percentage_to_fraction)
                .unwrap_or(0.0),
        }),
        "xfrm" => Some(BlipEffect::Transform {
            sx: node
                .attribute("sx")
                .and_then(crate::units::drawingml_percentage_to_fraction)
                .unwrap_or(1.0),
            sy: node
                .attribute("sy")
                .and_then(crate::units::drawingml_percentage_to_fraction)
                .unwrap_or(1.0),
            kx: node
                .attribute("kx")
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(0.0)
                / 60000.0,
            ky: node
                .attribute("ky")
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(0.0)
                / 60000.0,
            tx: node
                .attribute("tx")
                .and_then(crate::units::coordinate_to_emu)
                .unwrap_or(0) as f64,
            ty: node
                .attribute("ty")
                .and_then(crate::units::coordinate_to_emu)
                .unwrap_or(0) as f64,
        }),
        "alphaModFix" => node
            .attribute("amt")
            .map(crate::units::drawingml_percentage_to_fraction)
            .unwrap_or(Some(1.0))
            .map(|amt| BlipEffect::AlphaModFix { amt: amt.max(0.0) }),
        "alphaRepl" => node
            .attribute("a")
            .and_then(crate::units::drawingml_percentage_to_fraction)
            .map(|alpha| BlipEffect::AlphaRepl {
                alpha: alpha.clamp(0.0, 1.0),
            }),
        "clrRepl" => color(Some(node)).map(|(color, _)| BlipEffect::ColorReplace { color }),
        // extLst is metadata, not an authored raster effect.
        "extLst" => None,
        _ => None,
    }
}

/// §20.1.8.24 permits references to named effect containers. CT_Blip's
/// immediate chain does not necessarily contain the definition, so preserve
/// reached part-local recipes while XML and the theme resolver still exist.
/// No lookup occurs for ordinary or self-contained chains. Each missing name
/// needs a part search; unlike a whole-part retained index, this holds only
/// reached recipes (at most the existing 16-effect envelope). Element visits
/// and projected recipe nodes count against the same 2^27 image-effect-work
/// policy used in core, also accumulated by the package resource governor. These are
/// resource limits, not OOXML limits or compatibility guesses.
fn complete_named_effects(
    effects: &mut Vec<BlipEffect>,
    blip: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
    context: &mut EffectParseContext,
) {
    use std::collections::{BTreeSet, HashSet};
    fn collect(effects: &[BlipEffect], names: &mut HashSet<String>, refs: &mut BTreeSet<String>) {
        for effect in effects {
            match effect {
                BlipEffect::AlphaMod { container }
                | BlipEffect::Container { container }
                | BlipEffect::Blend { container, .. } => {
                    if let Some(name) = &container.name {
                        names.insert(name.clone());
                    }
                    collect(&container.effects, names, refs);
                }
                BlipEffect::Reference { reference }
                    if !["fill", "line", "fillLine", "children"].contains(&reference.as_str()) =>
                {
                    refs.insert(reference.clone());
                }
                BlipEffect::Definitions { containers, .. } => {
                    for container in containers {
                        if let Some(name) = &container.name {
                            names.insert(name.clone());
                        }
                        collect(&container.effects, names, refs);
                    }
                }
                _ => {}
            }
        }
    }
    let mut names = HashSet::new();
    let mut refs = BTreeSet::new();
    collect(effects, &mut names, &mut refs);
    let mut containers = Vec::new();
    let mut work: u64 = 0;
    if crate::resource::assert_healthy().is_err() {
        return;
    }
    while let Some(reference) = refs.pop_first() {
        if context.failure.is_some() {
            break;
        }
        if names.contains(&reference) {
            continue;
        }
        let before = work;
        let mut target = None;
        for candidate in blip.document().descendants() {
            work += 1;
            // Shared with core/image/pixel-budget.ts; stop the search itself,
            // not merely the later rendering, at the existing work ceiling.
            if work
                .saturating_add(context.work)
                .saturating_add(context.lookup_work)
                > MAX_IMAGE_EFFECT_LOOKUP_WORK
            {
                break;
            }
            if candidate.is_element()
                && matches!(candidate.tag_name().name(), "cont" | "effectDag")
                && candidate.attribute("name") == Some(reference.as_str())
            {
                target = Some(candidate);
                break;
            }
        }
        // Count the target projection before allocating strings, gradient
        // stops or nested fills. Both lookup and repeated retained projections
        // are package-accounted, preventing many pictures from expanding one
        // shared XML recipe quadratically. No work is done after poisoning.
        if let Some(target) = target {
            for _ in target.descendants() {
                work += 1;
                if work
                    .saturating_add(context.work)
                    .saturating_add(context.lookup_work)
                    > MAX_IMAGE_EFFECT_LOOKUP_WORK
                {
                    break;
                }
            }
        }
        if crate::resource::charge_image_effect_work(work - before, MAX_IMAGE_EFFECT_LOOKUP_WORK)
            .is_err()
            || work
                .saturating_add(context.work)
                .saturating_add(context.lookup_work)
                > MAX_IMAGE_EFFECT_LOOKUP_WORK
        {
            work = MAX_IMAGE_EFFECT_LOOKUP_WORK + 1;
            break;
        }
        if let Some(target) = target {
            let container = parse_container(target, resolver, tint_mode, context);
            if context.failure.is_some() {
                break;
            }
            names.insert(reference);
            collect(&container.effects, &mut names, &mut refs);
            containers.push(container);
            // One extra recipe is enough for the renderer's typed count quota;
            // do not expand an unbounded reference graph during parsing.
            if containers.len() > MAX_IMAGE_EFFECT_STEPS {
                break;
            }
        }
    }
    if work > 0 {
        context.lookup_work = context.lookup_work.saturating_add(work);
        effects.push(BlipEffect::Definitions { containers, work });
    }
}

/// §20.1.8.49 / §20.1.10.52 gives each preset as one or two outer
/// shadows. Values below are normative preset parameters, not measured tuning.
fn parse_preset_shadow(
    node: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
) -> Option<Vec<crate::effect::Shadow>> {
    let preset: u8 = node.attribute("prst")?.strip_prefix("shdw")?.parse().ok()?;
    if !(1..=20).contains(&preset) {
        return None;
    }
    let mut shadow = crate::effect::parse_shadow_effect(node, resolver, tint_mode);
    shadow.rot_with_shape = Some(false);
    match preset {
        3 => {
            shadow.algn = Some("b".into());
            shadow.ky = Some(40.89);
            shadow.sy = Some(0.5);
        }
        4 | 8 | 12 | 16 => {
            shadow.algn = Some("b".into());
            shadow.kx = Some(-40.89);
            shadow.sy = Some(if preset == 8 || preset == 16 {
                -0.5
            } else {
                0.5
            });
        }
        7 | 11 | 15 => {
            shadow.algn = Some("b".into());
            shadow.kx = Some(40.89);
            shadow.sy = Some(if preset == 7 || preset == 15 {
                -0.5
            } else {
                0.5
            });
        }
        9 => {
            shadow.algn = Some("tl".into());
            shadow.sx = Some(0.75);
            shadow.sy = Some(0.75);
        }
        10 => {
            shadow.algn = Some("br".into());
            shadow.sx = Some(1.25);
            shadow.sy = Some(1.25);
        }
        19 | 20 => {
            shadow.algn = Some("b".into());
            shadow.sy = Some(if preset == 19 { 0.5 } else { -1.0 });
        }
        _ => {}
    }
    let mut result = vec![shadow.clone()];
    if matches!(preset, 13 | 17 | 18) {
        let lighten = |i| {
            u8::from_str_radix(&shadow.color[i..i + 2], 16)
                .unwrap_or(0)
                .saturating_add(102)
        };
        shadow.color = format!("{:02X}{:02X}{:02X}", lighten(0), lighten(2), lighten(4));
        if preset == 13 {
            shadow.dist *= 2;
        } else {
            shadow.dir = (shadow.dir + 180.0) % 360.0;
        }
        result.push(shadow);
    }
    Some(result)
}

fn parse_container(
    node: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
    context: &mut EffectParseContext,
) -> EffectContainer {
    EffectContainer {
        kind: node.attribute("type").unwrap_or("sib").into(),
        name: node.attribute("name").map(str::to_owned),
        effects: parse_effect_children(node, resolver, tint_mode, context),
    }
}

fn parse_effect_fill(
    parent: Node<'_, '_>,
    resolver: &dyn crate::color::ThemeResolver,
    tint_mode: crate::color::TintMode,
    context: &mut EffectParseContext,
) -> Option<EffectFill> {
    let node = parent.children().find(|n| {
        n.is_element()
            && matches!(
                n.tag_name().name(),
                "noFill" | "solidFill" | "gradFill" | "pattFill" | "blipFill" | "grpFill"
            )
    })?;
    if !context.enter(node) {
        return None;
    }
    // Bound projection before allocating gradient stops, color transforms or
    // nested recipes. Descendants are iterative in roxmltree. Repeated inherited
    // expansion spends the same package-wide work budget as named lookup.
    for descendant in node.descendants() {
        if !context.charge(1) {
            context.active.remove(&node.id());
            return None;
        }
        // Do not let the tolerant shared brush parser substitute defaults for
        // malformed authored geometry. Defaults apply to absence only (§A.4.1).
        for attr in descendant.attributes() {
            let valid = match attr.name() {
                "l" | "t" | "r" | "b" | "pos" | "sx" | "sy" => {
                    crate::units::drawingml_percentage_to_fraction(attr.value())
                        .is_some_and(f64::is_finite)
                }
                "tx" | "ty" if descendant.tag_name().name() == "tile" => {
                    crate::units::coordinate_to_emu(attr.value()).is_some()
                }
                "ang" => attr.value().parse::<f64>().is_ok_and(f64::is_finite),
                "rotWithShape" | "scaled" => matches!(attr.value(), "true" | "false" | "0" | "1"),
                "dpi" => attr.value().parse::<u32>().is_ok(),
                _ => true,
            };
            if !valid {
                context.active.remove(&node.id());
                return None;
            }
        }
    }
    let fill = (|| {
        match node.tag_name().name() {
            "noFill" => Some(EffectFill::None),
            "grpFill" => {
                // §20.1.8.35 inherits the containing group's fill. Resolve it while
                // ancestor XML and the owning theme are available, rather than
                // asking a raster effect to guess parent paint after parsing.
                for ancestor in parent.ancestors().skip(1) {
                    if !context.charge(1) {
                        return None;
                    }
                    if let Some(properties) = ancestor
                        .children()
                        .find(|n| n.is_element() && n.tag_name().name() == "grpSpPr")
                    {
                        if properties == parent {
                            continue;
                        }
                        return parse_effect_fill(properties, resolver, tint_mode, context);
                    }
                }
                // No containing group supplied a fill: preserve an explicit
                // malformed-effect diagnostic instead of silently painting none.
                None
            }
            "solidFill" => crate::color::parse_color_node(node, resolver, tint_mode)
                .map(|color| EffectFill::Solid { color }),
            "gradFill" => {
                // CT_GradientStop@pos is a required fixed percentage (§20.1.8.36),
                // including Strict percentage lexemes. Shape parsing tolerates
                // bad stops; an effect recipe must reject the entire fill rather
                // than silently removing an authored color or defaulting its pos.
                let mut stops = 0u64;
                for stop in node
                    .descendants()
                    .filter(|n| n.is_element() && n.tag_name().name() == "gs")
                {
                    stops += 1;
                    if !stop
                        .attribute("pos")
                        .and_then(crate::units::drawingml_percentage_to_fraction)
                        .is_some_and(|p| p.is_finite() && (0.0..=1.0).contains(&p))
                        || crate::color::parse_color_node(stop, resolver, tint_mode).is_none()
                    {
                        return None;
                    }
                }
                // Sorting and validating stops precede allocation in the brush.
                // Charge this projection work as well as the XML visits above.
                if !context.charge(stops.saturating_mul(2 + u64::from((stops + 1).ilog2()))) {
                    return None;
                }
                crate::fill::parse_grad_fill(node, resolver, tint_mode).map(|g| {
                    EffectFill::Gradient {
                        stops: g.stops,
                        angle: g.angle,
                        grad_type: g.grad_type,
                        scaled: g.scaled,
                        path: g.path,
                        fill_to_rect: g.fill_to_rect,
                        tile_rect: g.tile_rect,
                        flip: g.flip,
                        rot_with_shape: g.rot_with_shape,
                    }
                })
            }
            "pattFill" => {
                // The shared painter's absent-field black/white/pct50 defaults
                // are library policy. An authored but unresolvable color is a
                // malformed recipe, never permission to substitute those defaults.
                for color in node
                    .children()
                    .filter(|n| n.is_element() && matches!(n.tag_name().name(), "fgClr" | "bgClr"))
                {
                    crate::color::parse_color_node(color, resolver, tint_mode)?;
                }
                let p = crate::fill::parse_patt_fill(node, resolver, tint_mode);
                Some(EffectFill::Pattern {
                    fg: p.fg,
                    bg: p.bg,
                    preset: p.preset,
                })
            }
            "blipFill" => {
                let blip = node
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "blip")?;
                let svg_image_path = svg_blip_rid(blip);
                let image_path = blip_embed_rid(&blip).or_else(|| svg_image_path.clone())?;
                let stretch = node
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "stretch");
                let tile = node
                    .children()
                    .find(|n| n.is_element() && n.tag_name().name() == "tile");
                let mut blip_effects = parse_effect_children(blip, resolver, tint_mode, context);
                complete_named_effects(&mut blip_effects, blip, resolver, tint_mode, context);
                Some(EffectFill::Image {
                    mime_type: mime_from_ext(&image_path).into(),
                    image_path,
                    svg_image_path,
                    src_rect: parse_src_rect(node),
                    fill_rect: stretch.and_then(crate::fill::parse_fill_rect),
                    tile: tile.map(crate::fill::parse_tile),
                    dpi: node.attribute("dpi").and_then(|v| v.parse().ok()),
                    stretch: stretch.is_some(),
                    rot_with_shape: node
                        .attribute("rotWithShape")
                        .map(|v| matches!(v, "true" | "1")),
                    blip_effects,
                })
            }
            _ => None,
        }
    })();
    context.active.remove(&node.id());
    fill
}

/// Resolve image relationships in nested fills with the owning part's map;
/// a relationship id is never sent to the renderer as an archive path.
pub fn resolve_blip_effect_images(
    effects: &mut [BlipEffect],
    resolve: &mut impl FnMut(&str) -> Option<String>,
) {
    for effect in effects {
        match effect {
            BlipEffect::Definitions { containers, .. } => {
                for container in containers {
                    resolve_blip_effect_images(&mut container.effects, resolve);
                }
            }
            BlipEffect::AlphaMod { container }
            | BlipEffect::Container { container }
            | BlipEffect::Blend { container, .. } => {
                resolve_blip_effect_images(&mut container.effects, resolve)
            }
            BlipEffect::FillOverlay {
                fill:
                    EffectFill::Image {
                        image_path,
                        mime_type,
                        svg_image_path,
                        blip_effects,
                        ..
                    },
                ..
            }
            | BlipEffect::Fill {
                fill:
                    EffectFill::Image {
                        image_path,
                        mime_type,
                        svg_image_path,
                        blip_effects,
                        ..
                    },
            } => {
                if let Some(path) = resolve(image_path) {
                    *mime_type = mime_from_ext(&path).into();
                    *image_path = path;
                }
                if let Some(svg) = svg_image_path.as_mut() {
                    if let Some(path) = resolve(svg) {
                        *svg = path;
                    }
                }
                resolve_blip_effect_images(blip_effects, resolve);
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const ASVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    // ISO/IEC 29500 Strict: a blip declaring the relationships namespace under
    // the `http://purl.oclc.org/ooxml/officeDocument/relationships` URI still
    // resolves its `r:embed` rId (the `relationships::STRICT` branch of
    // `blip_embed_rid`).
    #[test]
    fn blip_embed_rid_reads_strict_relationships_ns() {
        let r_strict = crate::ns::relationships::STRICT;
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{r_strict}" r:embed="rIdStrict"/>"#);
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            blip_embed_rid(&doc.root_element()).as_deref(),
            Some("rIdStrict")
        );
    }

    #[test]
    fn svg_blip_rid_reads_extension_regardless_of_prefix() {
        // The svgBlip uses a different prefix (asvg:) on purpose: matching is by
        // namespace-local name, so the prefix must not matter.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}" r:embed="rIdPng">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rIdPng"));
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn svg_blip_rid_requires_registered_uri_and_namespace() {
        for (uri, namespace) in [
            ("{00000000-0000-0000-0000-000000000000}", ASVG_NS),
            (SVG_BLIP_EXT_URI, "urn:example:not-microsoft-svg"),
        ] {
            let xml = format!(
                r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:x="{namespace}">
                     <a:extLst><a:ext uri="{uri}"><x:svgBlip r:embed="rIdWrong"/></a:ext></a:extLst>
                   </a:blip>"#
            );
            let doc = Document::parse(&xml).unwrap();
            assert_eq!(svg_blip_rid(doc.root_element()), None);
        }
    }

    #[test]
    fn parse_src_rect_reads_fractions_and_defaults_absent_edges() {
        // l/r given, t/b absent → fractions = raw/100000, absent ⇒ 0.
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/><a:srcRect l="32560" r="3829"/></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let sr = parse_src_rect(doc.root_element()).expect("srcRect present");
        assert!((sr.l - 0.3256).abs() < 1e-9);
        assert!((sr.r - 0.03829).abs() < 1e-9);
        assert_eq!(sr.t, 0.0);
        assert_eq!(sr.b, 0.0);

        let strict_xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/><a:srcRect l="25%" r="-10%"/></a:blipFill>"#
        );
        let strict = Document::parse(&strict_xml).unwrap();
        let sr = parse_src_rect(strict.root_element()).expect("Strict srcRect present");
        assert_eq!(sr.l, 0.25);
        assert_eq!(sr.r, -0.1);
    }

    #[test]
    fn parse_src_rect_none_when_absent_or_all_zero() {
        let none = format!(r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/></a:blipFill>"#);
        assert!(parse_src_rect(Document::parse(&none).unwrap().root_element()).is_none());
        let zero = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:srcRect l="0" t="0" r="0" b="0"/></a:blipFill>"#
        );
        assert!(parse_src_rect(Document::parse(&zero).unwrap().root_element()).is_none());
    }

    #[test]
    fn svg_blip_only_has_no_raster_embed() {
        // sample-12 shape: the blip has no r:embed at all, only the svgBlip ext.
        let xml = format!(
            r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:asvg="{ASVG_NS}">
                 <a:extLst>
                   <a:ext uri="{SVG_BLIP_EXT_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                 </a:extLst>
               </a:blip>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip), None);
        assert_eq!(svg_blip_rid(blip).as_deref(), Some("rIdSvg"));
    }

    #[test]
    fn plain_blip_has_no_svg() {
        let xml = format!(r#"<a:blip xmlns:a="{A_NS}" xmlns:r="{R_NS}" r:embed="rId1"/>"#);
        let doc = Document::parse(&xml).unwrap();
        let blip = doc.root_element();
        assert_eq!(blip_embed_rid(&blip).as_deref(), Some("rId1"));
        assert_eq!(svg_blip_rid(blip), None);
    }

    #[test]
    fn mime_table_covers_svg_and_common_rasters() {
        assert_eq!(mime_from_ext("../media/image2.svg"), "image/svg+xml");
        assert_eq!(mime_from_ext("a/b/image1.PNG"), "image/png");
        // Equivalent spellings of a stored part name (§6.2.2.3).
        assert_eq!(mime_from_ext("word/media/image1.%73vg"), "image/svg+xml");
        assert_eq!(
            mime_from_ext("word/media/image1.%53%56%47"),
            "image/svg+xml"
        );
        assert_eq!(mime_from_ext("x.jpeg"), "image/jpeg");
        assert_eq!(mime_from_ext("x.webp"), "image/webp");
        assert_eq!(mime_from_ext("x.tif"), "image/tiff");
        assert_eq!(mime_from_ext("x.TIFF"), "image/tiff");
        assert_eq!(mime_from_ext("noext"), "application/octet-stream");
    }

    #[test]
    fn mime_table_covers_metafiles() {
        // WMF/EMF blips (charts, diagrams) get the conventional metafile MIMEs,
        // not the application/octet-stream fallback, so the renderer's content
        // sniff has a sensible MIME to pair with each part.
        assert_eq!(mime_from_ext("word/media/image1.wmf"), "image/wmf");
        assert_eq!(mime_from_ext("ppt/media/image3.emf"), "image/emf");
        // Case-insensitive, like the raster entries.
        assert_eq!(mime_from_ext("a/b/CHART.WMF"), "image/wmf");
        assert_eq!(mime_from_ext("x.EMF"), "image/emf");
    }

    // A minimal ThemeResolver mapping a couple of scheme names, so the duotone
    // scheme-color path is exercised without a full parser.
    struct DuoResolver;
    impl crate::color::ThemeResolver for DuoResolver {
        fn resolve_scheme_color(&self, name: &str) -> Option<String> {
            match name {
                "accent1" => Some("4472C4".to_owned()),
                _ => None,
            }
        }
    }

    #[test]
    fn parse_blip_alpha_reads_amt_fraction() {
        // amt=70000 → 0.70 (the sample-9.xlsx alphaModFix).
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip r:embed="rId1" xmlns:r="{R_NS}"><a:alphaModFix amt="70000"/></a:blip></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let a = parse_blip_alpha(doc.root_element()).expect("alpha present");
        assert!((a - 0.70).abs() < 1e-9);

        let strict = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="70%"/></a:blip></a:blipFill>"#
        );
        let a = parse_blip_alpha(Document::parse(&strict).unwrap().root_element())
            .expect("Strict alpha present");
        assert!((a - 0.70).abs() < 1e-9);
    }

    #[test]
    fn parse_blip_alpha_none_when_absent_or_fully_opaque() {
        let none = format!(r#"<a:blipFill xmlns:a="{A_NS}"><a:blip/></a:blipFill>"#);
        assert!(parse_blip_alpha(Document::parse(&none).unwrap().root_element()).is_none());
        // amt=100000 (100%) ⇒ nothing to apply.
        let opaque = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="100000"/></a:blip></a:blipFill>"#
        );
        assert!(parse_blip_alpha(Document::parse(&opaque).unwrap().root_element()).is_none());
        let defaulted = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix/></a:blip></a:blipFill>"#
        );
        assert!(parse_blip_alpha(Document::parse(&defaulted).unwrap().root_element()).is_none());

        for amt in ["99990", "99999"] {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaModFix amt="{amt}"/></a:blip></a:blipFill>"#
            );
            assert!(
                parse_blip_alpha(Document::parse(&xml).unwrap().root_element()).is_some(),
                "a non-identity authored alpha must be retained: {amt}"
            );
        }
    }

    #[test]
    fn parse_blip_duotone_resolves_two_colours_with_transforms() {
        // The sample-9.xlsx effect: clr1=black (prstClr), clr2=srgbClr DAB6BA with
        // lumMod/lumOff/tint/satMod transforms. clr1 is the dark endpoint.
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}">
                 <a:blip r:embed="rId1" xmlns:r="{R_NS}"/>
                 <a:duotone>
                   <a:prstClr val="black"/>
                   <a:srgbClr val="DAB6BA">
                     <a:lumMod val="20000"/><a:lumOff val="80000"/>
                     <a:tint val="45000"/><a:satMod val="400000"/>
                   </a:srgbClr>
                 </a:duotone>
               </a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let duo = parse_blip_duotone(
            doc.root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        )
        .expect("duotone present");
        // clr1 = black.
        assert_eq!(duo.clr1, "000000");
        // clr2 is DAB6BA lightened by the transforms → a light pink (all channels
        // high, R the largest). We assert the shape, not an exact hex (the exact
        // value depends on the shared transform math, reported separately).
        let r = u8::from_str_radix(&duo.clr2[0..2], 16).unwrap();
        let g = u8::from_str_radix(&duo.clr2[2..4], 16).unwrap();
        let b = u8::from_str_radix(&duo.clr2[4..6], 16).unwrap();
        assert!(r > 200, "clr2 R should be high (light pink), got {r}");
        assert!(
            r >= g && r >= b,
            "clr2 should be pink (R largest): {}",
            duo.clr2
        );
    }

    #[test]
    fn parse_blip_effects_keeps_document_order_and_duotone_position() {
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}">
                 <a:blip r:embed="rId1" xmlns:r="{R_NS}">
                   <a:clrChange><a:clrFrom><a:srgbClr val="ffffff"/></a:clrFrom>
                     <a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>
                   <a:duotone><a:prstClr val="black"/><a:srgbClr val="FF0000"/></a:duotone>
                   <a:grayscl/><a:biLevel thresh="50000"/><a:alphaModFix amt="50000"/>
                 </a:blip>
               </a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let effects = parse_blip_effects(
            doc.root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        );
        assert_eq!(
            effects,
            vec![
                BlipEffect::ColorChange {
                    from: "FFFFFF".into(),
                    from_alpha: 1.0,
                    to: "FFFFFF".into(),
                    to_alpha: 0.0,
                    use_alpha: false,
                },
                BlipEffect::Duotone {
                    clr1: "000000".into(),
                    clr2: "FF0000".into()
                },
                BlipEffect::Grayscale,
                BlipEffect::BiLevel { thresh: 0.5 },
                BlipEffect::AlphaModFix { amt: 0.5 },
            ]
        );
        assert_eq!(
            serde_json::to_value(&effects[3]).unwrap(),
            serde_json::json!({"type": "biLevel", "thresh": 0.5})
        );
        // A duotone alone keeps the duotone-only path (no effect list).
        let duotone_only = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip xmlns:r="{R_NS}" r:embed="rId1"><a:duotone><a:prstClr val="black"/><a:srgbClr val="FF0000"/></a:duotone></a:blip></a:blipFill>"#
        );
        assert!(parse_blip_effects(
            Document::parse(&duotone_only).unwrap().root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear
        )
        .is_empty());
    }

    #[test]
    fn named_image_fill_cycles_and_deep_expansion_are_bounded_during_parsing() {
        let parse = |count: usize, cycle: bool| {
            let mut definitions = String::new();
            for index in 0..count {
                let next = if cycle { 0 } else { index + 1 };
                definitions.push_str(&format!(r#"<a:cont name="m{index}" type="tree"><a:fillOverlay blend="over"><a:blipFill>
                  <a:blip r:embed="rId1"><a:alphaMod><a:cont type="tree"><a:effect ref="m{next}"/></a:cont></a:alphaMod></a:blip>
                  <a:stretch/></a:blipFill></a:fillOverlay></a:cont>"#));
            }
            let xml = format!(
                r#"<a:root xmlns:a="{A_NS}" xmlns:r="{R_NS}">{definitions}
              <a:blipFill><a:blip><a:alphaMod><a:cont type="tree"><a:effect ref="m0"/></a:cont></a:alphaMod></a:blip></a:blipFill>
            </a:root>"#
            );
            let doc = Document::parse(&xml).unwrap();
            let node = doc
                .root_element()
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "blipFill")
                .unwrap();
            serde_json::to_string(&parse_blip_effects(
                node,
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear,
            ))
            .unwrap()
        };
        assert!(parse(1, true).contains("cyclic effect reference"));
        assert!(parse(MAX_IMAGE_EFFECT_STEPS + 1, false).contains("effectCountLimit"));
    }

    #[test]
    fn inherited_picture_fill_cycle_is_a_diagnostic_not_a_stack_overflow() {
        let xml = format!(
            r#"<a:grpSp xmlns:a="{A_NS}" xmlns:r="{R_NS}">
          <a:grpSpPr><a:blipFill><a:blip r:embed="rId1">
            <a:fillOverlay blend="over"><a:grpFill/></a:fillOverlay>
          </a:blip><a:stretch/></a:blipFill></a:grpSpPr>
          <a:pic><a:blipFill><a:blip r:embed="rId2">
            <a:fillOverlay blend="over"><a:grpFill/></a:fillOverlay>
          </a:blip></a:blipFill></a:pic></a:grpSp>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let child = doc
            .descendants()
            .find(|n| n.tag_name().name() == "pic")
            .unwrap()
            .children()
            .find(|n| n.is_element())
            .unwrap();
        let effects = parse_blip_effects(
            child,
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        );
        assert!(
            matches!(&effects[..], [BlipEffect::Unsupported { effect }] if effect.contains("cyclic"))
        );
    }

    #[test]
    fn every_inline_expansion_shares_depth_and_projection_limits() {
        for (open, close) in [
            ("<a:cont>", "</a:cont>"),
            ("<a:alphaMod><a:cont>", "</a:cont></a:alphaMod>"),
            ("<a:blend blend=\"over\"><a:cont>", "</a:cont></a:blend>"),
            (
                "<a:fillOverlay blend=\"over\"><a:blipFill><a:blip r:embed=\"rId1\">",
                "</a:blip><a:stretch/></a:blipFill></a:fillOverlay>",
            ),
        ] {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}" xmlns:r="{R_NS}"><a:blip>{}<a:grayscl/>{}</a:blip></a:blipFill>"#,
                open.repeat(40),
                close.repeat(40)
            );
            let doc = Document::parse(&xml).unwrap();
            let effects = parse_blip_effects(
                doc.root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear,
            );
            assert!(
                matches!(&effects[..], [BlipEffect::EffectCountLimit { observed }] if *observed > MAX_IMAGE_EFFECT_STEPS)
            );
        }
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip>{}</a:blip></a:blipFill>"#,
            "<a:grayscl/>".repeat(10000)
        );
        let doc = Document::parse(&xml).unwrap();
        assert!(matches!(
            &parse_blip_effects(
                doc.root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear
            )[..],
            [BlipEffect::EffectCountLimit { .. }]
        ));
    }

    #[test]
    fn references_retain_named_recipes_outside_the_immediate_blip_chain() {
        let xml = format!(
            r#"<a:root xmlns:a="{A_NS}">
          <a:effectDag name="external" type="tree"><a:alphaRepl a="50%"/></a:effectDag>
          <a:blipFill><a:blip><a:alphaMod><a:cont type="tree"><a:effect ref="external"/></a:cont></a:alphaMod></a:blip></a:blipFill>
        </a:root>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let node = doc
            .descendants()
            .find(|n| n.tag_name().name() == "blipFill")
            .unwrap();
        let effects =
            parse_blip_effects(node, &DuoResolver, crate::color::TintMode::PowerPointLinear);
        let BlipEffect::Definitions { containers, work } = &effects[1] else {
            panic!("missing reference definition");
        };
        assert!(*work > 0);
        assert_eq!(containers[0].name.as_deref(), Some("external"));
        assert_eq!(
            containers[0].effects,
            vec![BlipEffect::AlphaRepl { alpha: 0.5 }]
        );
    }

    #[test]
    fn recursive_fills_do_not_drop_bad_colors_or_default_bad_stop_positions() {
        let parse = |fill: &str| {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:fillOverlay blend="over">{fill}</a:fillOverlay></a:blip></a:blipFill>"#
            );
            let doc = Document::parse(&xml).unwrap();
            serde_json::to_value(parse_blip_effects(
                doc.root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear,
            ))
            .unwrap()
        };
        let valid_stop = r#"<a:gs pos="25%"><a:srgbClr val="FF0000"/></a:gs>"#;
        let valid = parse(&format!(
            r#"<a:gradFill><a:gsLst>{valid_stop}<a:gs pos="75%"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin/></a:gradFill>"#
        ));
        assert_eq!(valid[0]["fill"]["stops"][0]["position"], 0.25);
        assert_eq!(valid[0]["fill"]["stops"][1]["position"], 0.75);
        for fill in [
            format!(r#"<a:gradFill><a:gsLst>{valid_stop}<a:gs pos="100000"><a:schemeClr val="unknown"/></a:gs></a:gsLst></a:gradFill>"#),
            format!(r#"<a:gradFill><a:gsLst>{valid_stop}<a:gs pos="bad"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>"#),
            r#"<a:pattFill prst="horz"><a:fgClr><a:schemeClr val="unknown"/></a:fgClr></a:pattFill>"#.into(),
            format!(r#"<a:gradFill><a:gsLst>{valid_stop}</a:gsLst><a:lin ang="bad"/></a:gradFill>"#),
            "<a:grpFill/>".into(),
        ] {
            assert_eq!(parse(&fill)[0]["type"], "unsupported");
        }
    }

    #[test]
    fn recursive_fills_preserve_duotone_and_inherit_ancestor_group_fill() {
        let xml = format!(
            r#"<a:grpSp xmlns:a="{A_NS}" xmlns:r="{R_NS}">
          <a:grpSpPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:grpSpPr>
          <a:grpSp><a:grpSpPr><a:xfrm/><a:grpFill/></a:grpSpPr><a:pic><a:blipFill><a:blip>
            <a:fillOverlay blend="over"><a:grpFill/></a:fillOverlay>
            <a:fillOverlay blend="mult"><a:blipFill><a:blip r:embed="rId2">
              <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="00FF00"/></a:duotone>
            </a:blip><a:stretch/></a:blipFill></a:fillOverlay>
          </a:blip></a:blipFill></a:pic></a:grpSp></a:grpSp>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let blip_fill = doc
            .descendants()
            .find(|n| n.tag_name().name() == "blipFill")
            .unwrap();
        let effects = parse_blip_effects(
            blip_fill,
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        );
        let json = serde_json::to_value(effects).unwrap();
        assert_eq!(
            json[0]["fill"],
            serde_json::json!({"fillType":"solid", "color":"FF0000"})
        );
        assert_eq!(
            json[1]["fill"]["blipEffects"][0],
            serde_json::json!({"type":"duotone", "clr1":"000000", "clr2":"00FF00"})
        );
    }

    #[test]
    fn authored_effects_retain_parameters_recursive_masks_and_fill_colors() {
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}" xmlns:r="{R_NS}"><a:blip>
          <a:alphaMod><a:cont type="tree"><a:alphaRepl a="50%"/>
            <a:alphaMod><a:cont><a:alphaInv/></a:cont></a:alphaMod></a:cont></a:alphaMod>
          <a:blur rad="12700" grow="false"/><a:blur/>
          <a:fillOverlay blend="screen"><a:gradFill><a:gsLst>
            <a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="FFFFFF"><a:alpha val="50%"/></a:srgbClr></a:gs>
          </a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:fillOverlay>
          <a:hsl hue="7200000" sat="-50%" lum="25%"/><a:tint hue="10800000" amt="-25%"/>
          <a:alphaInv><a:srgbClr val="FF0000"/></a:alphaInv>
          <a:duotone><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr><a:srgbClr val="FFFFFF"/></a:duotone>
          <a:fillOverlay blend="mult"><a:blipFill><a:blip r:embed="rId2"/><a:stretch/></a:blipFill></a:fillOverlay>
        </a:blip></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let mut effects = parse_blip_effects(
            doc.root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        );
        resolve_blip_effect_images(&mut effects, &mut |rid| {
            (rid == "rId2").then(|| "media/overlay.png".into())
        });
        let json = serde_json::to_value(effects).unwrap();
        assert_eq!(json[0]["container"]["kind"], "tree");
        assert_eq!(
            json[0]["container"]["effects"][1]["container"]["effects"][0]["type"],
            "alphaInv"
        );
        assert_eq!(
            json[1],
            serde_json::json!({"type":"blur", "radius":12700.0, "grow":false})
        );
        assert_eq!(
            json[2],
            serde_json::json!({"type":"blur", "radius":0.0, "grow":true})
        );
        assert_eq!(json[3]["fill"]["stops"][1]["color"], "FFFFFF80");
        assert_eq!(json[3]["fill"]["angle"], 90.0);
        assert_eq!(
            json[4],
            serde_json::json!({"type":"hsl", "hue":120.0, "sat":-0.5, "lum":0.25})
        );
        assert_eq!(
            json[5],
            serde_json::json!({"type":"tint", "hue":180.0, "amt":-0.25})
        );
        assert_eq!(
            json[6],
            serde_json::json!({"type":"alphaInv", "color":"FF0000"})
        );
        assert_eq!(json[7]["clr1"], "00000000");
        assert_eq!(json[8]["fill"]["imagePath"], "media/overlay.png");
        assert_eq!(json[8]["fill"]["mimeType"], "image/png");
    }

    #[test]
    fn ordered_effects_own_alpha_palettes_and_replacement() {
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip>
            <a:alphaModFix amt="50000"/>
            <a:clrChange useA="1"><a:clrFrom><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:clrFrom>
              <a:clrTo><a:srgbClr val="FF0000"/></a:clrTo></a:clrChange>
            <a:lum bright="90000"/>
            <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FF0000"/></a:duotone>
            <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="00FF00"/></a:duotone>
            <a:clrRepl><a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr></a:clrRepl>
            </a:blip></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        let effects = parse_blip_effects(
            doc.root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear,
        );
        let json = serde_json::to_value(effects).unwrap();
        assert_eq!(
            json[0],
            serde_json::json!({"type":"alphaModFix", "amt":0.5})
        );
        assert_eq!(json[1]["fromAlpha"], 128.0 / 255.0);
        assert_eq!(
            json[3],
            serde_json::json!({"type":"duotone", "clr1":"000000", "clr2":"FF0000"})
        );
        assert_eq!(json[4]["clr2"], "00FF00");
        assert_eq!(
            json[5],
            serde_json::json!({"type":"colorReplace", "color":"FF0000"})
        );
    }

    #[test]
    fn alpha_effects_preserve_sequence_and_schema_units() {
        let xml = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip>
            <a:alphaBiLevel thresh="50%"/><a:alphaCeiling/><a:alphaFloor/>
            <a:alphaInv/><a:alphaModFix amt="200000"/><a:alphaRepl a="25%"/>
            </a:blip></a:blipFill>"#
        );
        let doc = Document::parse(&xml).unwrap();
        assert_eq!(
            parse_blip_effects(
                doc.root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear
            ),
            vec![
                BlipEffect::AlphaBiLevel { thresh: 0.5 },
                BlipEffect::AlphaCeiling,
                BlipEffect::AlphaFloor,
                BlipEffect::AlphaInv { color: None },
                BlipEffect::AlphaModFix { amt: 2.0 },
                BlipEffect::AlphaRepl { alpha: 0.25 },
            ]
        );
    }

    #[test]
    fn malformed_effects_remain_explicit_diagnostics() {
        for effect in [
            "alphaMod",
            "fillOverlay",
            "biLevel",
            "clrChange",
            "clrRepl",
            "duotone",
            "alphaRepl",
            "alphaBiLevel",
        ] {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:{effect}/></a:blip></a:blipFill>"#
            );
            let doc = Document::parse(&xml).unwrap();
            assert_eq!(
                parse_blip_effects(
                    doc.root_element(),
                    &DuoResolver,
                    crate::color::TintMode::PowerPointLinear
                ),
                vec![BlipEffect::Unsupported {
                    effect: effect.into()
                }]
            );
        }
    }

    #[test]
    fn sibling_containers_keep_identity_branches() {
        for identity in [
            r#"<a:alphaModFix amt="100000"/>"#,
            r#"<a:alphaModFix/>"#,
            r#"<a:lum bright="0" contrast="0"/>"#,
        ] {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip><a:alphaMod><a:cont type="sib">
                {identity}<a:alphaInv/></a:cont></a:alphaMod></a:blip></a:blipFill>"#
            );
            let doc = Document::parse(&xml).unwrap();
            let parsed = parse_blip_effects(
                doc.root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear,
            );
            let BlipEffect::AlphaMod { container } = &parsed[0] else {
                panic!("missing alphaMod")
            };
            let expected = if identity.contains("alphaModFix") {
                BlipEffect::AlphaModFix { amt: 1.0 }
            } else {
                BlipEffect::Luminance {
                    bright: 0.0,
                    contrast: 0.0,
                }
            };
            assert_eq!(
                container.effects,
                vec![expected, BlipEffect::AlphaInv { color: None }]
            );
        }
    }

    #[test]
    fn parse_blip_effects_reads_luminance_brightness_and_contrast() {
        let parse = |lum: &str| {
            let xml = format!(
                r#"<a:blipFill xmlns:a="{A_NS}"><a:blip xmlns:r="{R_NS}" r:embed="rId1">{lum}</a:blip></a:blipFill>"#
            );
            parse_blip_effects(
                Document::parse(&xml).unwrap().root_element(),
                &DuoResolver,
                crate::color::TintMode::PowerPointLinear,
            )
        };
        let effects = parse(r#"<a:lum bright="70000" contrast="-70000"/>"#);
        assert_eq!(
            effects,
            vec![BlipEffect::Luminance {
                bright: 0.7,
                contrast: -0.7
            }]
        );
        assert_eq!(
            serde_json::to_value(&effects[0]).unwrap(),
            serde_json::json!({"type": "luminance", "bright": 0.7, "contrast": -0.7})
        );
        assert_eq!(
            parse(r#"<a:lum bright="100000"/>"#),
            vec![BlipEffect::Luminance {
                bright: 1.0,
                contrast: 0.0
            }]
        );
        assert!(parse(r#"<a:alphaModFix/>"#).is_empty());
        assert!(parse(r#"<a:alphaModFix amt="100000"/>"#).is_empty());
        assert!(parse(r#"<a:lum/>"#).is_empty());
        assert!(parse(r#"<a:lum bright="0" contrast="0"/>"#).is_empty());
    }

    #[test]
    fn parse_blip_duotone_none_when_absent_or_single_colour() {
        let none = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:blip r:embed="rId1" xmlns:r="{R_NS}"/></a:blipFill>"#
        );
        assert!(parse_blip_duotone(
            Document::parse(&none).unwrap().root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear
        )
        .is_none());
        // Only one colour child ⇒ not a valid duotone; dropped.
        let one = format!(
            r#"<a:blipFill xmlns:a="{A_NS}"><a:duotone><a:srgbClr val="112233"/></a:duotone></a:blipFill>"#
        );
        assert!(parse_blip_duotone(
            Document::parse(&one).unwrap().root_element(),
            &DuoResolver,
            crate::color::TintMode::PowerPointLinear
        )
        .is_none());
    }
}
