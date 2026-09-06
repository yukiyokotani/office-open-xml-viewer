//! Serde output model types (the pptx JSON wire shape) + their serde skip
//! helpers and pure transform types. Extracted verbatim from `lib.rs`; `lib.rs`
//! re-exports these via `pub use types::*`.

use ooxml_common::blip::{Duotone, SrcRect};
use ooxml_common::drawing::{DrawingGroupSpec, DrawingGroupTransform, DrawingRect};
use ooxml_common::math::MathNode;
use ooxml_common::text::SpaceLine;
use serde::{Deserialize, Serialize};

// Chart data-model structs now live in `ooxml_common::chart` (the Rust mirror
// of core's TS `ChartModel`). The parser builds a `ChartModel` and emits it as
// the single nested `chart` field of `ChartElement` — the pptx JSON shape the
// TS renderer consumes without a per-field adapter. Both the legacy chart
// (`parse_chart_part`) and chartEx (`parse_chartex_part`) structure parses now
// live in `ooxml_common::chart`, so this crate only needs `ChartModel` as the
// nested payload of `ChartElement`; the pptx adapters delegate the rest.
pub(crate) use ooxml_common::chart::ChartModel;

/// A gradient color stop. The owned type + `<a:gs>` parse now live in
/// `ooxml_common::fill` (shared DrawingML fill grammar); re-exported here under
/// the former name so `Fill::Gradient { stops: Vec<GradStop> }` and its
/// byte-identical `{"position":..,"color":".."}` JSON are unchanged.
pub(crate) use ooxml_common::fill::GradStop;

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Presentation {
    pub(crate) slide_width: i64,
    pub(crate) slide_height: i64,
    pub(crate) slides: Vec<Slide>,
    /// Default text color from theme dk1 (hex 6 chars, e.g. "383838").
    pub(crate) default_text_color: Option<String>,
    /// Theme major (heading) font resolved name (e.g. "Aptos Display", "Nunito Sans").
    pub(crate) major_font: Option<String>,
    /// Theme minor (body) font resolved name (e.g. "Aptos", "Nunito Sans").
    pub(crate) minor_font: Option<String>,
    /// Theme hyperlink colour (hex 6 chars). Used by the renderer to colour
    /// hyperlink runs whose rPr does not specify an explicit colour.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) hlink_color: Option<String>,
    /// Theme followed-hyperlink colour (hex 6 chars). Reserved for visited-link
    /// styling — emitted so the renderer can colour visited hyperlinks once
    /// click history is wired up.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) fol_hlink_color: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Slide {
    pub(crate) index: usize,
    /// 1-based slide number (index + 1); used for slidenum field rendering
    pub(crate) slide_number: usize,
    /// The slide's normalized OPC part name (e.g. `ppt/slides/slide3.xml`),
    /// resolved through `presentation.xml.rels` in `sldIdLst` order. An internal
    /// hyperlink slide jump (`<a:hlinkClick action="ppaction://hlinksldjump">`,
    /// ECMA-376 §21.1.2.3.5) resolves its rel Target to this same part name, so
    /// the TS side can build a `partName → index` map and turn the click into a
    /// slide index. `None` (and omitted from JSON) only for a placeholder slide
    /// whose part path wasn't recorded; every healthy slide carries it.
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) part_name: Option<String>,
    pub(crate) background: Option<Fill>,
    pub(crate) elements: Vec<SlideElement>,
    /// Provenance for each rendered element, index-aligned with `elements`.
    pub(crate) element_sources: Vec<SlideElementSource>,
    /// `ppt/notesSlides/notesSlideN.xml` plain text — the speaker-notes pane
    /// content as a single string (paragraphs joined with '\n'). `None` when
    /// the slide has no notes part. Renderer ignores this; surfaced for tools.
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) notes: Option<String>,
    /// Classic ECMA-376 comments and modern threaded comments related from this
    /// slide. The part target is relationship-defined and need not use the
    /// conventional `ppt/comments/` directory.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub(crate) comments: Vec<PptxComment>,
    /// `<p:sld show="0">` — slide is hidden in the slide show (§19.3.1.38).
    /// Omitted from JSON when false so existing snapshots are unchanged.
    #[serde(skip_serializing_if = "std::ops::Not::not", default)]
    pub(crate) hidden: bool,
    /// RB7 partial degradation: when this slide's part (`ppt/slides/slideN.xml`
    /// or a dependency it needs) could not be parsed, the deck still opens with
    /// the OTHER slides intact and this one becomes a placeholder carrying the
    /// part-tagged error (e.g. `"ppt/slides/slide3.xml: <detail>"`). `None` (and
    /// omitted from JSON) for every healthy slide, so existing snapshots are
    /// byte-for-byte unchanged. The renderer paints a visible error placeholder.
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) parse_error: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub(crate) enum SlideElementOrigin {
    Master,
    Layout,
    Slide,
}

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SlideElementSource {
    pub(crate) origin: SlideElementOrigin,
}

/// Explicit target carried by a PowerPoint modern comment. [MS-PPTX]
/// §2.16.3.3 requires one anchor group; drawing and text anchors use the
/// content-moniker lists from [MS-ODRAWXML] §2.29.3.20/21.
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Eq)]
#[serde(tag = "type", rename_all = "camelCase")]
pub(crate) enum PptxCommentAnchor {
    Slide,
    DrawingElement {
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(rename = "elementId")]
        element_id: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(rename = "creationId")]
        creation_id: Option<String>,
    },
    TextRange {
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(rename = "elementId")]
        element_id: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        start: Option<i32>,
        #[serde(skip_serializing_if = "Option::is_none")]
        length: Option<i32>,
    },
    Unknown,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PptxComment {
    /// `<cm @authorId>` — author-list identifier. Required by ECMA-376
    /// §19.4.1, but optional here so a malformed comment can still degrade to
    /// its text instead of poisoning the slide.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) author_id: Option<u32>,
    /// Modern `p188:cm@authorId` ([MS-PPTX] §2.16). Kept separate from the
    /// classic decimal `authorId` so adding modern comments does not change the
    /// existing public field's type.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) modern_author_id: Option<String>,
    /// Modern `p188:cm@id`. Classic identity remains `(authorId, idx)`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    /// `<cm @idx>` — identifier unique among this author's comments.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) index: Option<u32>,
    /// Resolved author name from `ppt/commentAuthors.xml` (`<cmAuthor @id>`
    /// matches `<cm @authorId>`). `None` when authors file is missing or
    /// authorId is out of range.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) author: Option<String>,
    /// `<cm @dt>` — ISO-8601 date string when the comment was authored.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) date: Option<String>,
    /// `<p:pos @x/@y>` comment anchor on the slide surface, in EMU
    /// (ECMA-376 §19.4.5). Each coordinate remains independently optional for
    /// malformed input; the standard requires both.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) x: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) y: Option<i64>,
    /// Modern comment targets. Multiple drawing/text moniker lists mean one
    /// comment can explicitly target multiple slide elements.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub(crate) anchors: Vec<PptxCommentAnchor>,
    /// Modern `p188:cm@status`; omitted for classic comments. The schema
    /// defaults a missing value to `active`, so the parser resolves that value.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) status: Option<String>,
    /// Plain-text body from classic `<p:text>` or modern `<p188:txBody>`.
    pub(crate) text: String,
    /// Modern threaded replies in authored order.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub(crate) replies: Vec<PptxCommentReply>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PptxCommentReply {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) author_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) date: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) status: Option<String>,
    pub(crate) text: String,
}

// serde-facing parser output enum; the variant sizes follow the OOXML element
// model. Boxing the large Shape variant would change the JSON serialization
// shape only cosmetically while complicating 30+ construction/match sites, for
// no real benefit on this parse-once-then-serialize type.
#[allow(clippy::large_enum_variant)]
// `Clone` lets `build_master_bundle` pre-extract the master's decorative shapes
// once (per cached master) and hand each slide its own owned copy, instead of
// re-parsing the master XML and re-walking its spTree for every slide (D4). The
// derive is otherwise inert on this parse-once-then-serialize type.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub(crate) enum SlideElement {
    Shape(ShapeElement),
    Picture(PictureElement),
    Table(TableElement),
    Chart(ChartElement),
    Media(MediaElement),
}

/// A chart graphic frame on a slide. The chart payload itself is the shared
/// [`ChartModel`]; only the frame geometry (`x`/`y`/`width`/`height`, in EMU)
/// and the `type` discriminant are pptx-specific.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ChartElement {
    /// `<p:nvGraphicFramePr><p:cNvPr @id>` for the chart frame.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) rotation: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    pub(crate) chart: ChartModel,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableElement {
    /// `<p:nvGraphicFramePr><p:cNvPr @id>` for the table frame.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) rotation: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    /// Column widths in EMU
    pub(crate) cols: Vec<i64>,
    pub(crate) rows: Vec<TableRow>,
    /// `<a:tblPr rtl="1">` (ECMA-376 §21.1.3.13): right-to-left table —
    /// column 0 renders at the right edge. Skipped when false/absent.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub(crate) rtl: bool,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableRow {
    /// Row height in EMU
    pub(crate) height: i64,
    pub(crate) cells: Vec<TableCell>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TableCell {
    pub(crate) text_body: Option<TextBody>,
    pub(crate) fill: Option<Fill>,
    /// Whether `tcPr` authored a fill choice, including explicit `noFill`.
    /// Kept parser-internal so table styles cannot overwrite direct formatting.
    #[serde(skip)]
    pub(crate) has_direct_fill: bool,
    /// Default run text colour inherited from the table style (`<a:tcTxStyle>`),
    /// used when a run carries no explicit colour. Hex, no `#`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) text_color: Option<String>,
    pub(crate) border_l: Option<Stroke>,
    pub(crate) border_r: Option<Stroke>,
    pub(crate) border_t: Option<Stroke>,
    pub(crate) border_b: Option<Stroke>,
    /// Direct border presence is distinct from its rendered `Option<Stroke>`:
    /// `<a:lnL><a:noFill/></a:lnL>` is authored and must suppress a style line.
    #[serde(skip)]
    pub(crate) has_direct_border_l: bool,
    #[serde(skip)]
    pub(crate) has_direct_border_r: bool,
    #[serde(skip)]
    pub(crate) has_direct_border_t: bool,
    #[serde(skip)]
    pub(crate) has_direct_border_b: bool,
    #[serde(skip)]
    pub(crate) has_direct_diagonal_tl: bool,
    #[serde(skip)]
    pub(crate) has_direct_diagonal_tr: bool,
    /// Diagonal from top-left to bottom-right (tl2br)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) diagonal_tl: Option<Stroke>,
    /// Diagonal from top-right to bottom-left (tr2bl)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) diagonal_tr: Option<Stroke>,
    /// Column span (gridSpan attribute)
    pub(crate) grid_span: u32,
    /// Row span
    pub(crate) row_span: u32,
    /// Horizontal merge continuation (cell has no content, covered by left neighbour)
    pub(crate) h_merge: bool,
    /// Vertical merge continuation
    pub(crate) v_merge: bool,
}

/// Explicit text frame for a shape, sourced from a SmartArt drawing's
/// `<dsp:txXfrm>` (Microsoft diagram drawing extension). Coordinates are
/// absolute EMU in the same space as the shape's `x/y/width/height`, so the
/// group-transform / graphicFrame-offset passes adjust them in lock-step.
/// When present the renderer lays text out in this rectangle instead of the
/// preset/ellipse-derived text rectangle — PowerPoint stores the actual text
/// region here (e.g. an arrow's label sits past an overlapping circle node,
/// a roundRect's label avoids an overlapping bottom badge).
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextRect {
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
}

/// DrawingML 3D rotation in sphere coordinates — ECMA-376 §20.1.5.11
/// (`CT_SphereCoords`). All three angles are stored in **degrees** (the XML
/// carries 60000ths of a degree; we divide once here). Per the spec, `lat` and
/// `lon` are latitude/longitude and `rev` is the revolution about the resulting
/// view axis.
#[derive(Serialize, Deserialize, Debug, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Rot3d {
    /// Latitude — rotation about the horizontal (X) axis, degrees.
    pub(crate) lat: f64,
    /// Longitude — rotation about the vertical (Y) axis, degrees.
    pub(crate) lon: f64,
    /// Revolution — in-plane rotation about the view (Z) axis, degrees.
    pub(crate) rev: f64,
}

/// `<a:camera>` — ECMA-376 §20.1.5.5 (`CT_Camera`). Defines the camera that
/// views the 3D scene. `prst` selects one of the 62 preset cameras
/// (§20.1.10.47); `fov`/`zoom` optionally override the preset.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Camera3d {
    /// Preset camera name (`ST_PresetCameraType`), e.g. "perspectiveRelaxed".
    pub(crate) prst: String,
    /// Field-of-view override in **degrees** (60000ths in XML). None = use the
    /// preset's default FOV. Only meaningful for perspective presets.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) fov: Option<f64>,
    /// Zoom factor as a unit ratio (1.0 = 100%). XML carries an
    /// `ST_PositivePercentage` (e.g. 100000 = 100%); we divide by 100000.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) zoom: Option<f64>,
    /// Camera rotation override (`<a:rot>`). None = use the preset's base
    /// orientation unchanged.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) rot: Option<Rot3d>,
}

/// `<a:lightRig>` — ECMA-376 §20.1.5.9 (`CT_LightRig`). Parsed for Phase B
/// (lighting/bevel shading); the Phase A camera renderer ignores it.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct LightRig {
    /// Light-rig preset (`ST_LightRigType`), e.g. "threePt".
    pub(crate) rig: String,
    /// Light direction (`ST_LightRigDirection`): tl/t/tr/l/r/bl/b/br.
    pub(crate) dir: String,
    /// Optional rotation override of the rig.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) rot: Option<Rot3d>,
}

/// `<a:scene3d>` — ECMA-376 §20.1.4.1.41 (`CT_Scene3D`). Holds the camera and
/// light rig for a shape's 3D scene.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Scene3d {
    pub(crate) camera: Camera3d,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) light_rig: Option<LightRig>,
}

/// `<a:bevel>` — ECMA-376 §20.1.5.3 (`CT_Bevel`). Lengths in EMU; `w`/`h`
/// default to 76200 EMU and `prst` to "circle" per the schema.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Bevel3d {
    /// Bevel width in EMU.
    pub(crate) w: i64,
    /// Bevel height in EMU.
    pub(crate) h: i64,
    /// Bevel preset name (`ST_BevelPresetType`).
    pub(crate) prst: String,
}

/// `<a:sp3d>` — ECMA-376 §20.1.5.12 (`CT_Shape3D`). Parsed in full but **not
/// rendered in Phase A** (camera-only). The contour/extrusion/bevel surfaces
/// are Phase B; the renderer reads only `scene3d` for the perspective
/// projection.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Sp3d {
    /// Z position of the shape's front face in EMU (default 0).
    #[serde(skip_serializing_if = "is_zero_i64")]
    #[serde(default)]
    pub(crate) z: i64,
    /// Extrusion (depth) height in EMU (default 0).
    #[serde(skip_serializing_if = "is_zero_i64")]
    #[serde(default)]
    pub(crate) extrusion_h: i64,
    /// Contour (outline) width in EMU (default 0).
    #[serde(skip_serializing_if = "is_zero_i64")]
    #[serde(default)]
    pub(crate) contour_w: i64,
    /// Contour colour (`<a:contourClr>` child, ECMA-376 §20.1.5.12). Resolved
    /// hex (e.g. "969696"). `None` when absent (the schema default is to reuse
    /// the shape's line/fill colour, which the renderer does not approximate).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) contour_clr: Option<String>,
    /// Extrusion side-wall colour (`<a:extrusionClr>` child).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) extrusion_clr: Option<String>,
    /// Preset surface material (`ST_PresetMaterialType`), default "warmMatte".
    pub(crate) prst_material: String,
    /// Top bevel.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) bevel_t: Option<Bevel3d>,
    /// Bottom bevel.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) bevel_b: Option<Bevel3d>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ShapeElement {
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) rotation: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    /// OOXML preset name (e.g. "rect", "ellipse") or "custGeom" when custom paths are used.
    pub(crate) geometry: String,
    pub(crate) fill: Option<Fill>,
    pub(crate) stroke: Option<Stroke>,
    pub(crate) text_body: Option<TextBody>,
    /// Default text color from p:style > fontRef (hex). Overrides renderer default
    /// when present; individual run colors still take precedence.
    pub(crate) default_text_color: Option<String>,
    /// Custom geometry paths (only set when geometry == "custGeom").
    /// Outer vec: one entry per <a:path>; inner vec: path commands with coords in [0,1].
    pub(crate) cust_geom: Option<Vec<Vec<PathCmd>>>,
    /// First adjustment value from prstGeom avLst (e.g. trapezoid inset).
    /// Value is in OOXML units (0–100000 range).
    pub(crate) adj: Option<f64>,
    /// Second adjustment value from prstGeom avLst (e.g. arrow-head width).
    pub(crate) adj2: Option<f64>,
    /// Third adjustment value from prstGeom avLst (e.g. callout tip x).
    pub(crate) adj3: Option<f64>,
    /// Fourth adjustment value from prstGeom avLst (e.g. callout tip y).
    pub(crate) adj4: Option<f64>,
    /// Fifth-through-eighth adjustment values (needed by callouts like
    /// accentBorderCallout3 whose polyline uses up to 8 adj values).
    pub(crate) adj5: Option<f64>,
    pub(crate) adj6: Option<f64>,
    pub(crate) adj7: Option<f64>,
    pub(crate) adj8: Option<f64>,
    /// Drop shadow from spPr > effectLst > outerShdw (None if not present).
    pub(crate) shadow: Option<Shadow>,
    /// Inner (inset) shadow from spPr > effectLst > innerShdw.
    /// ECMA-376 §20.1.8.21 (CT_InnerShadowEffect).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) inner_shadow: Option<Shadow>,
    /// Coloured glow halo from spPr > effectLst > glow.
    /// ECMA-376 §20.1.8.17 (CT_GlowEffect).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) glow: Option<Glow>,
    /// Soft (feathered) edge from spPr > effectLst > softEdge.
    /// ECMA-376 §20.1.8.31 (CT_SoftEdgesEffect).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) soft_edge: Option<SoftEdge>,
    /// Mirrored reflection from spPr > effectLst > reflection.
    /// ECMA-376 §20.1.8.27 (CT_ReflectionEffect).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) reflection: Option<Reflection>,
    /// `<p:nvSpPr><p:cNvPr @id>` — DrawingML cNvPr `id` attribute. Stable
    /// per-slide identifier surfaced for tools that need to reference a shape
    /// (MCP, scripted edits). Renderer ignores it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    /// `<p:nvSpPr><p:cNvPr @name>` — author-visible name (e.g. "Title 1",
    /// "Rectangle 5"). Useful for tools that want a human-readable handle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) name: Option<String>,
    /// Shape-level hyperlink resolved from `<p:nvSpPr><p:cNvPr><a:hlinkClick @r:id>`
    /// via slide _rels (ECMA-376 §21.1.2.3.5). For an EXTERNAL link this is the
    /// URL; for an INTERNAL slide-jump it is the resolved internal part name.
    /// None when the shape has no hlinkClick. Clicking anywhere on the shape
    /// activates it (distinct from a text-run hyperlink).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) hyperlink: Option<String>,
    /// Raw `<a:hlinkClick @action>` on the shape's cNvPr (e.g.
    /// "ppaction://hlinksldjump"), marking an internal action. See
    /// `TextRunData::hyperlink_action`. None when absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) hyperlink_action: Option<String>,
    /// `<p:nvSpPr><p:nvPr><p:ph @type>` — placeholder semantic type
    /// ("title" / "ctrTitle" / "body" / "subTitle" / "ftr" / "sldNum" /
    /// "dt" / "obj" / "pic" / etc., ECMA-376 §19.7.10). `None` for
    /// non-placeholder shapes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) placeholder_type: Option<String>,
    /// `<p:ph @idx>` — placeholder index used by the slide-layout chain to
    /// disambiguate multiple body / picture placeholders on a layout.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) placeholder_idx: Option<u32>,
    /// Explicit text rectangle from a SmartArt `<dsp:txXfrm>`. `None` for
    /// ordinary shapes (renderer falls back to the preset text rectangle).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) text_rect: Option<TextRect>,
    /// `<p:spPr><a:scene3d>` (ECMA-376 §20.1.4.1.41 / §20.1.5.5) — 3D camera
    /// scene. When the camera is non-identity the renderer projects the shape's
    /// 2D drawing through the camera homography (Phase A).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) scene3d: Option<Scene3d>,
    /// `<p:spPr><a:sp3d>` (ECMA-376 §20.1.5.12) — 3D shape properties
    /// (bevel/contour/extrusion). Parsed but not rendered in Phase A.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sp3d: Option<Sp3d>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PictureElement {
    /// The source drawing node's `<p:cNvPr @id>`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) rotation: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    /// Embedded zip path of the raster image from the blip's own `r:embed`
    /// (e.g. "ppt/media/image1.png"). The renderer fetches the bytes lazily by
    /// path (see `extract_image`) instead of inlining base64. When the picture
    /// is a pure SVG with no raster `r:embed` (only the svgBlip extension
    /// below), this falls back to the SVG part's path so the element is always
    /// drawable; `mime_type` is then `image/svg+xml` and `svg_image_path` holds
    /// the same path.
    pub(crate) image_path: String,
    /// MIME type of the blip at `image_path` (e.g. `image/png`), derived from
    /// the part extension via `ooxml_common::blip::mime_from_ext`.
    pub(crate) mime_type: String,
    /// Microsoft 2016 SVG extension (`<a:blip><a:extLst><a:ext
    /// uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed>`):
    /// the `r:embed` points at the `.svg` part that is the *original* vector
    /// image, while `image_path` (the blip's own `r:embed`) is the PNG fallback
    /// PowerPoint rasterizes for compatibility. The zip path of that `.svg`
    /// part; the renderer prefers the vector original and falls back to the
    /// raster on a decode failure. None when the picture has no svgBlip
    /// extension (the common case). Its MIME is always `image/svg+xml` and is
    /// owned by the SVG decoder.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) svg_image_path: Option<String>,
    /// Intrinsic pixel width of the raster blip, read from the PNG IHDR at parse
    /// time. None for non-PNG payloads (unchanged from the prior
    /// `png_size_from_data_url` semantics). Consumed by the ink-fallback sizing.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) intrinsic_width_px: Option<u32>,
    /// Intrinsic pixel height of the raster blip (PNG IHDR). None for non-PNG.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) intrinsic_height_px: Option<u32>,
    /// Border line from `<p:pic><p:spPr><a:ln>` (ECMA-376 §20.1.2.2.24
    /// CT_LineProperties; §19.3.1.37 routes a `p:pic`'s spPr through
    /// CT_ShapeProperties, so a picture carries the same line as a shape). Same
    /// model as `ShapeElement.stroke`. `None` when there is no `<a:ln>` or it
    /// resolves to `<a:noFill/>` (border explicitly suppressed).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) stroke: Option<Stroke>,
    /// `<p:spPr><a:prstGeom prst="…">` preset name (e.g. "roundRect",
    /// "ellipse"). ECMA-376 §20.1.9.18: a picture's preset geometry is its clip
    /// silhouette and the path its border / contour hug. None = plain rectangle
    /// (prst="rect" or no prstGeom). custGeom takes priority when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) prst_geom: Option<String>,
    /// Adjust guides from the prstGeom `<a:avLst>` (1/1000-of-a-percent OOXML
    /// units, in `gd@name` declaration order). Index 0 = adj/adj1, 1 = adj2, …
    /// Empty / None → the preset's own declared defaults apply. Carried generically
    /// so any of the 186 presets (not just roundRect) can be reconstructed.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) prst_adjust: Option<Vec<i64>>,
    /// ECMA-376 §20.1.8.55 a:srcRect — source image crop in 1/100000 fractions of
    /// source width/height. Only serialized when any edge is non-zero.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) src_rect: Option<SrcRect>,
    /// a:blip > a:alphaModFix@amt (0.0–1.0). None = fully opaque.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) alpha: Option<f64>,
    /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour effect, resolved to its two
    /// endpoint colours (through the slide's theme palette). `None` = no duotone
    /// (the common case). When set, the renderer remaps the decoded raster along
    /// the `clr1`→`clr2` luminance ramp at decode time (see the shared core
    /// `applyDuotone`), matching PowerPoint's recolour.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) duotone: Option<Duotone>,
    /// `<p:spPr><a:custGeom>` — custom geometry path used as a clip on the
    /// blitted image. Same shape model as `ShapeElement.cust_geom` (one or more
    /// `<a:path>` whose coordinates are normalized into [0,1] of the bbox).
    /// None when the picture is a plain rectangle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) cust_geom: Option<Vec<Vec<PathCmd>>>,
    /// Drop shadow from spPr > effectLst > outerShdw. ECMA-376 §20.1.8.45.
    /// `p:spPr` is CT_ShapeProperties for pictures too (§19.3.1.37), so the
    /// same effects a shape carries apply to images.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) shadow: Option<Shadow>,
    /// Inner (inset) shadow from spPr > effectLst > innerShdw. §20.1.8.40.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) inner_shadow: Option<Shadow>,
    /// Coloured glow halo from spPr > effectLst > glow. §20.1.8.32.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) glow: Option<Glow>,
    /// Soft (feathered) edge from spPr > effectLst > softEdge. §20.1.8.53.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) soft_edge: Option<SoftEdge>,
    /// Mirrored reflection from spPr > effectLst > reflection. §20.1.8.50.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) reflection: Option<Reflection>,
    /// `<p:spPr><a:scene3d>` (ECMA-376 §20.1.4.1.41 / §20.1.5.5). A `p:pic`'s
    /// `spPr` is `CT_ShapeProperties` (§19.3.1.37), so 3D scenes apply to images
    /// too. When non-identity, the renderer projects the picture through the
    /// camera homography (Phase A).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) scene3d: Option<Scene3d>,
    /// `<p:spPr><a:sp3d>` (ECMA-376 §20.1.5.12). Parsed but not rendered in
    /// Phase A.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sp3d: Option<Sp3d>,
}

/// ECMA-376 §19.3.1.17/18 a:audioFile / a:videoFile and the
/// p14:media extension (embed attribute).
/// Represents a p:pic that acts as an audio/video placeholder.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct MediaElement {
    /// `<p:nvPicPr><p:cNvPr @id>` for the media frame.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) id: Option<String>,
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) rotation: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    /// "audio" or "video"
    pub(crate) media_kind: String,
    /// Zip path of the poster image (e.g. "ppt/media/image2.png"). Empty when
    /// the media element has no blipFill poster. Fetched lazily through the
    /// same getMedia API as `media_path` so large posters don't bloat the
    /// parse output's JSON.
    pub(crate) poster_path: String,
    /// Poster image MIME type (derived from extension). Empty when no poster.
    pub(crate) poster_mime_type: String,
    /// Path inside the pptx zip (e.g. "ppt/media/media2.mp4"). The renderer
    /// uses this with a separate getMedia API to pull bytes lazily, avoiding
    /// the cost of base64-encoding large videos into the parse result.
    pub(crate) media_path: String,
    /// Media MIME type (e.g. "audio/mpeg", "video/mp4")
    pub(crate) mime_type: String,
}

pub(crate) use ooxml_common::fill::{FillRect, TileInfo};

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Shadow {
    /// hex color (6 chars)
    pub(crate) color: String,
    /// opacity 0.0–1.0
    pub(crate) alpha: f64,
    /// blur radius in EMU
    pub(crate) blur: i64,
    /// distance from shape in EMU
    pub(crate) dist: i64,
    /// direction in degrees, clockwise from East
    pub(crate) dir: f64,
    /// Horizontal scale (1.0 = unchanged). Outer shadows only.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sx: Option<f64>,
    /// Vertical scale (1.0 = unchanged). Outer shadows only.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) sy: Option<f64>,
    /// Horizontal skew in degrees. Outer shadows only.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) kx: Option<f64>,
    /// Vertical skew in degrees. Outer shadows only.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) ky: Option<f64>,
    /// Alignment origin for scale/skew.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) algn: Option<String>,
    /// Whether the shadow transform follows shape rotation. Schema default true.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) rot_with_shape: Option<bool>,
}

/// ECMA-376 §20.1.8.17 (CT_GlowEffect) — coloured halo with blur radius.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Glow {
    /// hex color (6 chars)
    pub(crate) color: String,
    /// opacity 0.0–1.0
    pub(crate) alpha: f64,
    /// blur radius in EMU
    pub(crate) radius: i64,
}

/// ECMA-376 §20.1.8.31 (CT_SoftEdgesEffect) — feathers the shape's alpha
/// edge by `rad` EMU. The effect itself has no colour child; it consumes
/// the shape's existing fill / stroke alpha at the perimeter.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SoftEdge {
    /// Feather radius in EMU.
    pub(crate) radius: i64,
}

/// ECMA-376 §20.1.8.27 (CT_ReflectionEffect) — mirrored copy below the shape
/// with a linear alpha gradient. The full spec exposes 14 attributes; this
/// model carries the ones that meaningfully change the visual: blur radius,
/// distance, direction, the start/end alpha+position pair, and per-axis
/// scale. Unsupported attributes (kx/ky skew, algn, fadeDir, rotWithShape)
/// fall back to their spec defaults at render time.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Reflection {
    /// Blur radius in EMU.
    pub(crate) blur: i64,
    /// Offset distance from the shape, EMU.
    pub(crate) dist: i64,
    /// Direction in degrees, clockwise from East.
    pub(crate) dir: f64,
    /// Start alpha (0–1). Top of the gradient.
    pub(crate) st_a: f64,
    /// Start position along the gradient (0–1).
    pub(crate) st_pos: f64,
    /// End alpha (0–1).
    pub(crate) end_a: f64,
    /// End position along the gradient (0–1).
    pub(crate) end_pos: f64,
    /// Horizontal scale (1.0 = same width). Negative flips horizontally.
    pub(crate) sx: f64,
    /// Vertical scale (-1.0 default for a true mirror).
    pub(crate) sy: f64,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "fillType", rename_all = "camelCase")]
pub(crate) enum Fill {
    Solid {
        color: String,
    },
    None,
    #[serde(rename_all = "camelCase")]
    Gradient {
        stops: Vec<GradStop>,
        /// degrees, 0 = left→right, 90 = top→bottom
        angle: f64,
        /// "linear" | "radial"
        grad_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        scaled: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_to_rect: Option<FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        tile_rect: Option<FillRect>,
        #[serde(skip_serializing_if = "Option::is_none")]
        flip: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
    },
    /// Preset pattern fill — ECMA-376 §20.1.8.40 / §20.1.10.59 (ST_PresetPatternVal).
    #[serde(rename_all = "camelCase")]
    Pattern {
        /// Foreground colour (hex). Used for the "1" pixels of the pattern bitmap.
        fg: String,
        /// Background colour (hex). Used for the "0" pixels.
        bg: String,
        /// Preset value: pct5/pct10/.../horz/vert/cross/diagCross/lgGrid/smGrid etc.
        preset: String,
    },
    /// Image fill — ECMA-376 §20.1.8.14 `a:blipFill`. The referenced blip is
    /// resolved to its embedded zip path + mime at parse time; the renderer
    /// fetches the bytes lazily by path (see `extract_image`) instead of
    /// inlining base64. Both fill-modes are modelled and mutually exclusive:
    /// `stretch` (§20.1.8.56) carries a `fill_rect`; `tile` (§20.1.8.58)
    /// carries a `tile` descriptor (see `parse_blip_fill`).
    #[serde(rename_all = "camelCase")]
    Image {
        /// Embedded zip path of the blip (e.g. "ppt/media/image1.png").
        image_path: String,
        /// MIME type of the blip at `image_path` (e.g. `image/png`).
        mime_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        svg_image_path: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        dpi: Option<u32>,
        #[serde(skip_serializing_if = "Option::is_none")]
        rot_with_shape: Option<bool>,
        /// `<a:srcRect>` (§20.1.8.55) source-image crop. Unlike `fill_rect`,
        /// this applies before either stretch or tile placement.
        #[serde(skip_serializing_if = "Option::is_none")]
        src_rect: Option<SrcRect>,
        /// `<a:stretch><a:fillRect>` (§20.1.8.30 CT_RelativeRect). Edge insets
        /// as fractions of the fill region; negative values overscan past the
        /// bounding box. `None` when stretch has no fillRect (= full box) or
        /// the fill-mode is `tile`.
        #[serde(skip_serializing_if = "Option::is_none")]
        fill_rect: Option<FillRect>,
        #[serde(skip_serializing_if = "std::ops::Not::not")]
        stretch: bool,
        /// `<a:tile>` (§20.1.8.58). `Some` only when the blipFill is tiled;
        /// mutually exclusive with `fill_rect`.
        #[serde(skip_serializing_if = "Option::is_none")]
        tile: Option<TileInfo>,
        /// `a:blip > a:alphaModFix@amt` as a fraction (0.0–1.0). None = opaque.
        #[serde(skip_serializing_if = "Option::is_none")]
        alpha: Option<f64>,
        /// ECMA-376 §20.1.8.23 `<a:duotone>` recolour effect, resolved to its two
        /// endpoint colours through the slide theme (PowerPoint linear tint).
        /// `None` = no duotone. When present the renderer maps the blip's
        /// luminance ramp between the two colours (core `applyDuotone`), matching
        /// how a `<p:pic>` duotone recolours — a shared blip effect that a
        /// picture FILL (§20.1.8.14) may carry just as a picture element can.
        #[serde(skip_serializing_if = "Option::is_none")]
        duotone: Option<Duotone>,
    },
}

/// Arrow end descriptor for headEnd / tailEnd on a line.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ArrowEnd {
    /// OOXML type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow"
    #[serde(rename = "type")]
    pub(crate) kind: String,
    /// Width multiplier: "sm" | "med" | "lg"
    pub(crate) w: String,
    /// Length multiplier: "sm" | "med" | "lg"
    pub(crate) len: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct StrokeDashSegment {
    /// Dash length as a multiplier of the rendered line width.
    pub(crate) dash: f64,
    /// Gap length as a multiplier of the rendered line width.
    pub(crate) space: f64,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Stroke {
    pub(crate) color: String,
    pub(crate) width: i64,
    /// Non-solid line paint. Solid lines keep using `color` to preserve the
    /// compact historical wire shape; gradient/pattern lines retain their
    /// authored fill here.
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) fill: Option<Fill>,
    /// OOXML prstDash value: "dash", "dot", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDot", etc.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) dash_style: Option<String>,
    /// DrawingML custom dash pairs, normalized from percentage to line-width
    /// multipliers. Present values take precedence over `dash_style`.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub(crate) custom_dash: Vec<StrokeDashSegment>,
    /// DrawingML cap normalized to the Canvas vocabulary.
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) line_cap: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) line_join: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) miter_limit: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none", default)]
    pub(crate) alignment: Option<String>,
    /// Arrow at the start of the line (headEnd)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) head_end: Option<ArrowEnd>,
    /// Arrow at the end of the line (tailEnd)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) tail_end: Option<ArrowEnd>,
    /// ECMA-376 §20.1.8.42 ST_CompoundLine — "sng" (default) | "dbl" |
    /// "thinThick" | "thickThin" | "tri". None = single line.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) cmpd: Option<String>,
}

/// A single path command inside a custGeom pathLst.
/// Coordinates are normalised to [0.0, 1.0] relative to the path's w/h,
/// so the renderer can map them directly to shape-local pixel coordinates.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "cmd", rename_all = "camelCase")]
pub(crate) enum PathCmd {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    /// Cubic Bézier: two control points + endpoint
    CubicBezTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    /// Quadratic Bézier: one control point plus endpoint (§20.1.9.21).
    QuadBezTo {
        x1: f64,
        y1: f64,
        x: f64,
        y: f64,
    },
    /// Elliptical arc (all angles in degrees)
    ///
    /// The enum-level `#[serde(tag = ..., rename_all = "camelCase")]` renames the
    /// variant *tag* (`ArcTo` → `arcTo`) but NOT the variant's struct fields, so
    /// a per-variant `rename_all` is required for the multi-word fields. Without
    /// it the JSON carried `st_ang`/`sw_ang`, while the TS `PathCmd`
    /// (core/src/types/common.ts) reads `stAng`/`swAng` — the mismatch left the
    /// angles `undefined`, producing `NaN` coordinates and a missing arc. This
    /// mirrors the per-variant `rename_all` already used by `Fill`, `Bullet`,
    /// and `TextRun`.
    #[serde(rename_all = "camelCase")]
    ArcTo {
        wr: f64,
        hr: f64,
        st_ang: f64,
        sw_ang: f64,
    },
    Close,
}

/// `<a:bodyPr><a:prstTxWarp prst="…">` — WordArt text warp (ECMA-376
/// §20.1.9.19, `ST_TextShapeType`). The preset name (e.g. `"textArchUp"`,
/// `"textWave1"`) selects one of the 40 `presetTextWarpDefinitions.xml`
/// envelopes; `adj` carries the `<a:avLst>` guide values (thousandths of a
/// percent, OOXML convention) in `adj1`, `adj2`, … order. An empty `adj` means
/// the preset's declared defaults apply. Absent (`None`) → no warp.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextWarp {
    /// The `prst` attribute, e.g. `"textArchUp"`. Renderer lower-cases it to key
    /// the shared warp-preset table.
    pub(crate) preset: String,
    /// `<a:avLst>` adjust values in declaration order (adj1, adj2, …). Empty when
    /// the author supplied no `<a:gd>` overrides, so the preset defaults apply.
    #[serde(skip_serializing_if = "Vec::is_empty", default)]
    pub(crate) adj: Vec<i64>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextBody {
    pub(crate) vertical_anchor: String,
    pub(crate) paragraphs: Vec<Paragraph>,
    pub(crate) default_font_size: Option<f64>,
    /// Inherited bold from layout/master lstStyle defRPr (None = not inherited)
    pub(crate) default_bold: Option<bool>,
    /// Inherited italic from layout/master lstStyle defRPr (None = not inherited)
    pub(crate) default_italic: Option<bool>,
    /// Text insets in EMU. Defaults: lIns=rIns=91440, tIns=bIns=45720
    pub(crate) l_ins: i64,
    pub(crate) r_ins: i64,
    pub(crate) t_ins: i64,
    pub(crate) b_ins: i64,
    /// Whether text wraps inside the bounding box ("square") or not ("none")
    pub(crate) wrap: String,
    /// Text direction from bodyPr vert attribute: "horz" | "vert" | "vert270" | "eaVert" etc.
    pub(crate) vert: String,
    /// Auto-fit mode from bodyPr: "sp" = spAutoFit (shape grows), "norm" = normAutoFit (font shrinks), "none" = noAutofit
    pub(crate) auto_fit: String,
    /// `<a:normAutofit fontScale>` (ECMA-376 §21.1.2.1.3) — PowerPoint's stored
    /// pre-computed font-shrink ratio as a fraction (62500 → 0.625). None when
    /// absent; the renderer then re-derives the scale itself.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) font_scale: Option<f64>,
    /// `<a:normAutofit lnSpcReduction>` — stored line-spacing reduction fraction
    /// (20000 → 0.20). None when absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) ln_spc_reduction: Option<f64>,
    /// `<a:bodyPr numCol>` (ECMA-376 §20.1.10.34 / §21.1.2.1.1) — number of
    /// text columns inside the shape. Default 1. PowerPoint distributes
    /// paragraphs across columns left-to-right, top-to-bottom.
    #[serde(skip_serializing_if = "is_one")]
    #[serde(default = "one_u32")]
    pub(crate) num_col: u32,
    /// `<a:bodyPr spcCol>` — gap between columns in EMU. Default 0.
    /// Only meaningful when `num_col > 1`.
    #[serde(skip_serializing_if = "is_zero_i64")]
    #[serde(default)]
    pub(crate) spc_col: i64,
    /// `<a:bodyPr rtlCol>` (ECMA-376 §21.1.2.1.1) — when true the columns of a
    /// multi-column text body are laid out right-to-left. Default false. Only
    /// meaningful when `num_col > 1`.
    #[serde(skip_serializing_if = "is_false")]
    #[serde(default)]
    pub(crate) rtl_col: bool,
    /// `<a:bodyPr><a:prstTxWarp>` — WordArt text warp (ECMA-376 §20.1.9.19).
    /// None when the body has no warp (the common case), so existing text bodies
    /// serialize byte-identically. When present the renderer maps each glyph
    /// through the preset's envelope instead of laying the text out flat.
    #[serde(skip_serializing_if = "Option::is_none")]
    #[serde(default)]
    pub(crate) text_warp: Option<TextWarp>,
}

pub(crate) fn one_u32() -> u32 {
    1
}

pub(crate) fn is_one(n: &u32) -> bool {
    *n == 1
}

pub(crate) fn is_zero_i64(n: &i64) -> bool {
    *n == 0
}

pub(crate) fn is_false(b: &bool) -> bool {
    !*b
}

/// Bullet / list-item marker for a paragraph
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub(crate) enum Bullet {
    /// Explicitly no bullet (buNone)
    None,
    /// No bullet element present – inherit from layout/master
    Inherit,
    /// Character bullet (buChar)
    #[serde(rename_all = "camelCase")]
    Char {
        #[serde(rename = "char")]
        ch: String,
        color: Option<String>,
        /// Size as % of text size (100.0 = same size)
        size_pct: Option<f64>,
        /// `<a:buSzPts>` (ECMA-376 §21.1.2.4.10) — an ABSOLUTE marker size in
        /// points (18.0 = 18pt), independent of the run size. Mutually exclusive
        /// with `size_pct` (both are the one `EG_TextBulletSize` choice). Omitted
        /// from the JSON when absent (serde skip), so existing bullets that carry
        /// no `buSzPts` serialize byte-identically.
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(default)]
        size_pts: Option<f64>,
        font_family: Option<String>,
    },
    /// Auto-numbered bullet (buAutoNum)
    #[serde(rename_all = "camelCase")]
    AutoNum {
        num_type: String,
        start_at: Option<u32>,
        /// `<a:buClr>` (ECMA-376 §21.1.2.4.4) — the marker colour. `None` means
        /// no explicit buClr, so the renderer falls back to the buClrTx default
        /// (§21.1.2.4.5 — the first run's colour). Always serialized (like the
        /// `Char` variant's `color`) so the TS side sees a stable `color` key.
        color: Option<String>,
        /// `<a:buSzPct>` as a percentage of the first run's text size.
        size_pct: Option<f64>,
        /// `<a:buSzPts>` as an absolute marker size in points.
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(default)]
        size_pts: Option<f64>,
        /// Explicit `<a:buFont typeface>`; None means follow the first run.
        font_family: Option<String>,
    },
    /// Picture bullet (buBlip) — ECMA-376 §21.1.2.4.2 `<a:buBlip><a:blip
    /// r:embed="rIdN"/></a:buBlip>`. The `r:embed` is resolved to the blip's
    /// embedded **zip path** (e.g. "ppt/media/image1.png") + mime at parse time,
    /// exactly like `Fill::Image`; the renderer fetches the bytes lazily by path.
    /// `size_pct` carries `<a:buSzPct>` (§21.1.2.4.9) as a fraction of 100 (e.g.
    /// 80.0 = 80% of the text size); None means the spec default of 100%.
    #[serde(rename_all = "camelCase")]
    Blip {
        /// Embedded zip path of the bullet image (e.g. "ppt/media/image1.png").
        image_path: String,
        /// MIME type of the blip at `image_path` (e.g. `image/png`).
        mime_type: String,
        /// `<a:buSzPct val>` as a percentage (100.0 = same size as text). None
        /// when no explicit `<a:buSzPct>` is present (renderer uses 100%).
        size_pct: Option<f64>,
        /// `<a:buSzPts val>` (ECMA-376 §21.1.2.4.10) — an ABSOLUTE marker size in
        /// points, independent of the text size. Mutually exclusive with
        /// `size_pct`. Omitted from the JSON when absent (serde skip).
        #[serde(skip_serializing_if = "Option::is_none")]
        #[serde(default)]
        size_pts: Option<f64>,
    },
}

/// A tab stop defined in a paragraph's pPr > tabLst.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TabStop {
    /// Position in EMU from the left edge of the text area (after lIns)
    pub(crate) pos: i64,
    /// Alignment: "l" | "r" | "ctr" | "dec"
    pub(crate) algn: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Paragraph {
    pub(crate) alignment: String,
    /// Left margin in EMU
    pub(crate) mar_l: i64,
    /// Right margin in EMU
    pub(crate) mar_r: i64,
    /// First-line indent in EMU (negative = hanging indent for bullets)
    pub(crate) indent: i64,
    pub(crate) space_before: Option<i64>,
    pub(crate) space_after: Option<i64>,
    pub(crate) space_line: Option<SpaceLine>,
    /// List nesting level (0–8)
    pub(crate) lvl: u32,
    pub(crate) bullet: Bullet,
    /// Paragraph-level default run properties (from pPr > defRPr)
    pub(crate) def_font_size: Option<f64>,
    pub(crate) def_color: Option<String>,
    pub(crate) def_bold: Option<bool>,
    pub(crate) def_italic: Option<bool>,
    pub(crate) def_font_family: Option<String>,
    /// Tab stops from pPr > tabLst
    pub(crate) tab_stops: Vec<TabStop>,
    /// `<a:pPr defTabSz>` (ECMA-376 §21.1.2.2.7) — default tab interval in EMU.
    /// A `\t` with no reachable explicit stop advances to the next multiple of
    /// this grid (issue #1006). None when the paragraph carried no explicit
    /// value; the renderer then applies PowerPoint's 1-inch (914400 EMU) default.
    #[serde(skip_serializing_if = "Option::is_none")]
    #[serde(default)]
    pub(crate) def_tab_sz: Option<i64>,
    /// ECMA-376 §21.1.2.2.7 `<a:pPr rtl="1">` — right-to-left paragraph.
    /// When true and no explicit `algn`, the default alignment flips from
    /// "l" to "r". Carried through so the renderer can also flow runs RTL
    /// when bidi shaping is added.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub(crate) rtl: bool,
    /// ECMA-376 §21.1.2.2.7 `<a:pPr eaLnBrk>` — whether an East Asian word may
    /// be broken at a line wrap. xsd:boolean, default true when the attribute is
    /// omitted. true → CJK may break at character boundaries (kinsoku rules);
    /// false → an East Asian word must NOT be split mid-character. Resolved
    /// through the paragraph → body/list-style → layout/master cascade, mirroring
    /// `alignment`, so the renderer receives the effective value.
    pub(crate) ea_ln_brk: bool,
    pub(crate) runs: Vec<TextRun>,
}

// serde-facing parser output enum; same rationale as SlideElement — the Text
// variant is the common case and boxing it would add an allocation per run with
// no meaningful gain on this parse-once-then-serialize type.
#[allow(clippy::large_enum_variant)]
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type", rename_all = "camelCase")]
pub(crate) enum TextRun {
    Text(TextRunData),
    Break,
    /// An OMML equation embedded in the paragraph (ECMA-376 §22.1). PowerPoint
    /// stores these as `a14:m` inside `mc:AlternateContent`. `display` is true
    /// for `m:oMathPara` (block) math, false for inline `m:oMath`.
    #[serde(rename_all = "camelCase")]
    Math {
        nodes: Vec<MathNode>,
        display: bool,
        /// Paragraph default run size (pt) if declared; None → renderer inherits.
        #[serde(skip_serializing_if = "Option::is_none")]
        font_size: Option<f64>,
        /// Equation run colour (hex, no '#') from the math run's rPr solidFill;
        /// None → renderer uses the paragraph/body default colour.
        #[serde(skip_serializing_if = "Option::is_none")]
        color: Option<String>,
    },
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextRunData {
    pub(crate) text: String,
    /// None = not set (inherit from paragraph/body/layout defaults); Some(true/false) = explicit
    pub(crate) bold: Option<bool>,
    /// None = not set; Some(true/false) = explicit
    pub(crate) italic: Option<bool>,
    pub(crate) underline: bool,
    /// OOXML rPr @u value when explicit and != "sng" — e.g. "dbl", "dotted",
    /// "dash", "wavy", "heavy", "dotDash", … None means either no underline
    /// or the default single-line style (rPr @u = "sng" or unset truthy).
    /// ECMA-376 §21.1.2.3.9 (`rPr@u`); ST_TextUnderlineType §20.1.10.82.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) underline_style: Option<String>,
    /// Underline-specific colour from rPr > uFill > solidFill. None means the
    /// underline follows the text colour (uFillTx behaviour, the default).
    /// ECMA-376 §21.1.2.3.12 (CT_TextUnderlineFillGroupWrapper).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) underline_color: Option<String>,
    /// true when strike == "sngStrike" or "dblStrike"
    pub(crate) strikethrough: bool,
    /// true only when strike == "dblStrike" (renderer draws two parallel lines)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub(crate) strike_double: bool,
    pub(crate) font_size: Option<f64>,
    pub(crate) color: Option<String>,
    pub(crate) font_family: Option<String>,
    /// East Asian font family from rPr > ea (resolved through the theme).
    /// Renderer uses this for CJK runs. None = inherit from latin font.
    /// ECMA-376 §21.1.2.3.3 (CT_TextFont, ea variant).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) font_family_ea: Option<String>,
    /// Symbol font family from rPr > sym (resolved through the theme).
    /// Renderer uses this for symbol-range PUA glyphs (U+F0xx).
    /// ECMA-376 §21.1.2.3.10 (CT_TextFont, sym variant).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) font_family_sym: Option<String>,
    /// Baseline shift in thousandths of a point. Positive = superscript, negative = subscript.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) baseline: Option<i32>,
    /// Capitalisation from `rPr@cap` — ECMA-376 §21.1.2.3.9;
    /// ST_TextCapsType §20.1.10.64.
    /// "none" | "small" | "all". None = inherit / no transform.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) caps: Option<String>,
    /// Letter spacing in points after normalizing DrawingML `rPr@spc` from
    /// unitless hundredths of a point or `ST_UniversalMeasure` (ECMA-376
    /// §21.1.2.3.9; ST_TextPoint §20.1.10.74). Positive = looser, negative = tighter.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) letter_spacing: Option<f64>,
    /// Set for OOXML field elements (e.g. "slidenum" for slide number fields)
    pub(crate) field_type: Option<String>,
    /// Hyperlink target resolved from rPr > hlinkClick @r:id via slide _rels.
    /// For an EXTERNAL link this is the URL; for an INTERNAL slide-jump it is the
    /// resolved internal part name (e.g. "../slides/slide3.xml", TargetMode=Internal).
    /// None for runs without a:hlinkClick. ECMA-376 §21.1.2.3.5 (CT_Hyperlink).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) hyperlink: Option<String>,
    /// Raw `<a:hlinkClick @action>` string (e.g. "ppaction://hlinksldjump")
    /// when present. Its presence marks the link as an INTERNAL PowerPoint
    /// action (slide jump / first / last / ...) rather than an external URL;
    /// `hyperlink` then holds the resolved internal part name for a slide jump.
    /// None when the hlinkClick has no @action. ECMA-376 §21.1.2.3.5.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) hyperlink_action: Option<String>,
    /// ECMA-376 §20.1.8.45 (CT_OuterShadowEffect) — drop shadow on this run's
    /// glyphs from `<a:rPr><a:effectLst><a:outerShdw>`. Distinct from the
    /// shape-level shadow on `spPr`. None = no shadow on the run.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) shadow: Option<Shadow>,
    /// ECMA-376 §20.1.8.50 (CT_ReflectionEffect) — mirrored reflection of this
    /// run's glyphs from `<a:rPr><a:effectLst><a:reflection>`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) reflection: Option<Reflection>,
    /// ECMA-376 §20.1.2.2.24 (CT_TextOutlineEffect) — text glyph outline from
    /// `<a:rPr><a:ln w="EMU"><a:solidFill>...`. None = no outline; renderer
    /// just fillText. When set the renderer also strokeText with the given
    /// width (EMU) and colour.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) outline: Option<TextOutline>,
    /// ECMA-376 §21.1.2.3.4 — text highlight (marker) colour from
    /// `<a:rPr><a:highlight>`. The DrawingML highlight is a full CT_Color (any
    /// srgbClr / schemeClr / sysClr / prstClr + transforms), NOT the fixed
    /// 16-name enum WordprocessingML uses — so it resolves through the same
    /// colour pipeline as solidFill. Resolved hex without `#` (6-char opaque,
    /// or 8-char RRGGBBAA when an alpha transform applies). None = no
    /// highlight; renderer draws no background box.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) highlight: Option<String>,
}

/// Run-level text outline (`<a:rPr><a:ln>`). The width is the OOXML EMU
/// value (`w` attribute, 12700 EMU = 1 pt); the renderer converts to px.
#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct TextOutline {
    /// Outline width in EMU.
    pub(crate) width: i64,
    /// Resolved hex colour (no `#`). None = inherit from text fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) color: Option<String>,
}

#[derive(Clone, Debug, Default, Serialize)]
pub(crate) struct Transform {
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) cx: i64,
    pub(crate) cy: i64,
    /// Degrees, clockwise
    pub(crate) rot: f64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
}

#[derive(Clone, Debug, Default)]
pub(crate) struct GroupTransform {
    pub(crate) x: i64,
    pub(crate) y: i64,
    pub(crate) cx: i64,
    pub(crate) cy: i64,
    pub(crate) ch_x: i64,
    pub(crate) ch_y: i64,
    pub(crate) ch_cx: i64,
    pub(crate) ch_cy: i64,
    pub(crate) flip_h: bool,
    pub(crate) flip_v: bool,
    /// Group rotation in degrees, clockwise
    pub(crate) rot: f64,
}

impl GroupTransform {
    fn apply_to_transform(&self, t: Transform) -> Transform {
        let mapped = DrawingGroupTransform::from_group(DrawingGroupSpec {
            off_x: self.x as f64,
            off_y: self.y as f64,
            ext_x: self.cx as f64,
            ext_y: self.cy as f64,
            child_off_x: self.ch_x as f64,
            child_off_y: self.ch_y as f64,
            child_ext_x: self.ch_cx as f64,
            child_ext_y: self.ch_cy as f64,
            rotation_degrees: self.rot,
            flip_h: self.flip_h,
            flip_v: self.flip_v,
        })
        .apply_rect(DrawingRect {
            x: t.x as f64,
            y: t.y as f64,
            width: t.cx as f64,
            height: t.cy as f64,
            rotation_degrees: t.rot,
            flip_h: t.flip_h,
            flip_v: t.flip_v,
        });
        Transform {
            x: mapped.x.round() as i64,
            y: mapped.y.round() as i64,
            cx: mapped.width.round() as i64,
            cy: mapped.height.round() as i64,
            rot: mapped.rotation_degrees,
            flip_h: mapped.flip_h,
            flip_v: mapped.flip_v,
        }
    }
}

pub(crate) fn apply_group_transform_to_element(el: &mut SlideElement, gt: &GroupTransform) {
    match el {
        SlideElement::Shape(s) => {
            let t = Transform {
                x: s.x,
                y: s.y,
                cx: s.width,
                cy: s.height,
                rot: s.rotation,
                flip_h: s.flip_h,
                flip_v: s.flip_v,
            };
            let nt = gt.apply_to_transform(t);
            s.x = nt.x;
            s.y = nt.y;
            s.width = nt.cx;
            s.height = nt.cy;
            s.rotation = nt.rot;
            s.flip_h = nt.flip_h;
            s.flip_v = nt.flip_v;
            // Transform the explicit text frame in lock-step. It is axis-aligned
            // in local coords; pass rot=0/flip=false so it only translates+scales
            // (SmartArt drawings that carry txXfrm are not nested in rotated groups).
            if let Some(tr) = &mut s.text_rect {
                let tt = Transform {
                    x: tr.x,
                    y: tr.y,
                    cx: tr.width,
                    cy: tr.height,
                    rot: 0.0,
                    flip_h: false,
                    flip_v: false,
                };
                let ntt = gt.apply_to_transform(tt);
                tr.x = ntt.x;
                tr.y = ntt.y;
                tr.width = ntt.cx;
                tr.height = ntt.cy;
            }
        }
        SlideElement::Picture(p) => {
            let t = Transform {
                x: p.x,
                y: p.y,
                cx: p.width,
                cy: p.height,
                rot: p.rotation,
                flip_h: p.flip_h,
                flip_v: p.flip_v,
            };
            let nt = gt.apply_to_transform(t);
            p.x = nt.x;
            p.y = nt.y;
            p.width = nt.cx;
            p.height = nt.cy;
            p.rotation = nt.rot;
            p.flip_h = nt.flip_h;
            p.flip_v = nt.flip_v;
        }
        SlideElement::Table(tbl) => {
            // If the table has no xfrm (zero dimensions), it fills the group's child space.
            let (ex, ey, ecx, ecy) = if tbl.width == 0 && tbl.height == 0 {
                (gt.ch_x, gt.ch_y, gt.ch_cx, gt.ch_cy)
            } else {
                (tbl.x, tbl.y, tbl.width, tbl.height)
            };
            let t = Transform {
                x: ex,
                y: ey,
                cx: ecx,
                cy: ecy,
                rot: tbl.rotation,
                flip_h: tbl.flip_h,
                flip_v: tbl.flip_v,
            };
            let nt = gt.apply_to_transform(t);
            tbl.x = nt.x;
            tbl.y = nt.y;
            tbl.width = nt.cx;
            tbl.height = nt.cy;
            tbl.rotation = nt.rot;
            tbl.flip_h = nt.flip_h;
            tbl.flip_v = nt.flip_v;
        }
        SlideElement::Chart(chart) => {
            // If the chart graphicFrame has no xfrm (zero dimensions), it fills the group's child space.
            let (ex, ey, ecx, ecy) = if chart.width == 0 && chart.height == 0 {
                (gt.ch_x, gt.ch_y, gt.ch_cx, gt.ch_cy)
            } else {
                (chart.x, chart.y, chart.width, chart.height)
            };
            let t = Transform {
                x: ex,
                y: ey,
                cx: ecx,
                cy: ecy,
                rot: chart.rotation,
                flip_h: chart.flip_h,
                flip_v: chart.flip_v,
            };
            let nt = gt.apply_to_transform(t);
            chart.x = nt.x;
            chart.y = nt.y;
            chart.width = nt.cx;
            chart.height = nt.cy;
            chart.rotation = nt.rot;
            chart.flip_h = nt.flip_h;
            chart.flip_v = nt.flip_v;
        }
        SlideElement::Media(m) => {
            let t = Transform {
                x: m.x,
                y: m.y,
                cx: m.width,
                cy: m.height,
                rot: m.rotation,
                flip_h: m.flip_h,
                flip_v: m.flip_v,
            };
            let nt = gt.apply_to_transform(t);
            m.x = nt.x;
            m.y = nt.y;
            m.width = nt.cx;
            m.height = nt.cy;
            m.rotation = nt.rot;
            m.flip_h = nt.flip_h;
            m.flip_v = nt.flip_v;
        }
    }
}

#[cfg(test)]
mod group_transform_tests {
    use super::*;

    #[test]
    fn quarter_turned_child_uses_shared_non_uniform_group_mapping() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 127_000,
            cy: 254_000,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 127_000,
            ch_cy: 127_000,
            ..Default::default()
        };
        let mapped = group.apply_to_transform(Transform {
            x: 0,
            y: 50_800,
            cx: 127_000,
            cy: 25_400,
            rot: 90.0,
            ..Default::default()
        });

        assert_eq!((mapped.x, mapped.y), (0, 101_600));
        assert_eq!((mapped.cx, mapped.cy), (127_000, 50_800));
        assert_eq!(mapped.rot, 90.0);
    }

    #[test]
    fn rotated_flipped_group_uses_shared_annex_l_mapping() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 200,
            cy: 100,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 100,
            ch_cy: 100,
            rot: 90.0,
            flip_h: true,
            flip_v: false,
        };
        let mapped = group.apply_to_transform(Transform {
            x: 10,
            y: 20,
            cx: 20,
            cy: 10,
            rot: 15.0,
            flip_h: false,
            flip_v: true,
        });

        assert_eq!(
            (mapped.x, mapped.y, mapped.cx, mapped.cy),
            (105, 105, 40, 10)
        );
        assert_eq!(mapped.rot, 75.0);
        assert!(mapped.flip_h);
        assert!(mapped.flip_v);
    }

    #[test]
    fn group_transform_preserves_table_chart_and_media_frame_orientation() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 200,
            cy: 100,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 100,
            ch_cy: 100,
            rot: 90.0,
            flip_h: true,
            flip_v: false,
        };

        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let chart = crate::chart::parse_legacy_chart(xml, &std::collections::HashMap::new())
            .expect("minimal chart");
        let mut elements = vec![
            SlideElement::Table(TableElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                cols: vec![],
                rows: vec![],
                rtl: false,
            }),
            SlideElement::Chart(ChartElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                ..chart
            }),
            SlideElement::Media(MediaElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                media_kind: "video".into(),
                poster_path: String::new(),
                poster_mime_type: String::new(),
                media_path: String::new(),
                mime_type: "video/mp4".into(),
            }),
        ];

        for element in &mut elements {
            apply_group_transform_to_element(element, &group);
            match element {
                SlideElement::Table(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                SlideElement::Chart(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                SlideElement::Media(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                _ => unreachable!(),
            }
        }
    }
}
