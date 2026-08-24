// ===== Shared types re-exported from @silurus/ooxml-core =====
export type {
  PathCmd,
  Fill, SolidFill, NoFill, GradientFill, GradientStop, ImageFill, FillRect, SrcRect, TileInfo,
  Shadow,
  Glow,
  SoftEdge,
  Reflection,
  Stroke,
  SpaceLine,
  TabStop,
  TextRun, TextRunData, LineBreak,
  RenderOptions,
  ChartModel, ChartPlotGroup, ChartPlotGroupAxisSlot, ChartPlotGroupKind,
  ChartSeries, SecondaryValueAxis,
} from '@silurus/ooxml-core';

// ===== Presentation data model =====
// All positions and sizes are in EMUs (English Metric Units).
// 914400 EMU = 1 inch, 12700 EMU = 1 pt

import type { Bullet as CoreBullet, Fill, Stroke, TextBody as CoreTextBody, Paragraph as CoreParagraph, Shadow, Glow, SoftEdge, Reflection, PathCmd, ChartModel, Duotone } from '@silurus/ooxml-core';

/**
 * Picture bullet — ECMA-376 §21.1.2.4.2 `<a:buBlip><a:blip r:embed>`. The
 * embed is resolved to the blip's embedded zip path + mime at parse time
 * (mirrors {@link ImageFill}); the renderer fetches the bytes lazily by path
 * via the same `getCachedBitmap(imagePath, mimeType, fetchImage)` path used for
 * `pic`/blipFill. PPTX-only: the shared core `Bullet` union (used by docx/xlsx,
 * which have no picture bullets) does not carry this variant, so it lives on
 * the PPTX side, exactly like {@link Paragraph} extends the core paragraph.
 */
export interface BlipBullet {
  type: 'blip';
  /** Embedded zip path of the bullet image (e.g. "ppt/media/image1.png"). */
  imagePath: string;
  /** MIME type of the blip at {@link BlipBullet.imagePath} (e.g. `image/png`). */
  mimeType: string;
  /**
   * `<a:buSzPct val>` (ECMA-376 §21.1.2.4.9) as a percentage of the text size
   * (100 = same size). `null` when no explicit `<a:buSzPct>` is present, in
   * which case the renderer uses the spec default of 100%.
   */
  sizePct: number | null;
  /**
   * `<a:buSzPts val>` (ECMA-376 §21.1.2.4.10) — an ABSOLUTE marker size in
   * points, independent of the text size. Mutually exclusive with {@link
   * BlipBullet.sizePct} (both are the one `EG_TextBulletSize` choice). Omitted
   * when no explicit `<a:buSzPts>` is present; when present it takes precedence.
   */
  sizePts?: number;
}

/**
 * PPTX bullet marker. The shared core {@link CoreBullet} union
 * (none/inherit/char/autoNum) plus the PPTX-only picture bullet
 * ({@link BlipBullet}, §21.1.2.4.2). The parser emits the `blip` variant with
 * `type: "blip"`, so this is a discriminated union just like the core one.
 */
export type Bullet = CoreBullet | BlipBullet;

/**
 * PPTX paragraph. Extends the shared core `Paragraph` with the PPTX-only
 * `eaLnBrk` flag that the pptx parser emits but the shared core model does not
 * carry (docx/xlsx paragraphs don't surface it). Mirrors the Rust
 * `Paragraph` struct's `ea_ln_brk` field 1:1.
 *
 * Note on `bullet`: the parser also emits the picture-bullet variant
 * ({@link BlipBullet}, `type: "blip"`) at runtime, but `bullet` keeps the
 * narrower core type here because a TS interface can only *narrow* an inherited
 * property, not widen its union (the core `Paragraph.bullet` is used by
 * docx/xlsx, which have no picture bullets). Consumers that need the picture
 * variant narrow `bullet` with {@link asBullet} / a `type === 'blip'` check.
 */
export interface Paragraph extends CoreParagraph {
  /**
   * `<a:pPr eaLnBrk>` (ECMA-376 §21.1.2.2.7, xsd:boolean, default true). When
   * true, East Asian text may break at character boundaries (kinsoku rules);
   * when false, an East Asian word must not be split mid-character. The parser
   * resolves the paragraph → body/list-style → layout/master cascade and always
   * emits an effective boolean.
   */
  eaLnBrk: boolean;
  /**
   * `<a:pPr defTabSz>` (ECMA-376 §21.1.2.2.7) — the default tab-stop interval in
   * EMU. When a `\t` has no reachable explicit `a:tabLst` stop, it advances to
   * the next multiple of this grid (issue #1006). Absent ⇒ the renderer uses the
   * PowerPoint universal default of 914400 EMU (1 inch). Omitted from JSON when
   * the parser found no explicit value.
   */
  defTabSz?: number;
}

/**
 * View a paragraph's `bullet` as the PPTX {@link Bullet} union (which includes
 * the picture bullet). The pptx parser emits `{ type: "blip", … }` at runtime,
 * but the statically-typed `CoreParagraph.bullet` only lists none/inherit/char/
 * autoNum because that type is shared with docx/xlsx. This is the single, named
 * cast site so the widening is documented rather than scattered.
 */
export function asBullet(bullet: CoreBullet): Bullet {
  return bullet as Bullet;
}

/**
 * PPTX text body. Extends the shared core `TextBody` with PPTX-only bodyPr
 * fields that the pptx parser surfaces but the shared core model does not yet
 * carry.
 */
export interface TextBody extends CoreTextBody {
  /**
   * `<a:bodyPr rtlCol>` (ECMA-376 §21.1.2.1.1) — when true the columns of a
   * multi-column text body are laid out right-to-left. Defaults to false;
   * omitted from JSON when false. Only meaningful when `numCol > 1`.
   */
  rtlCol?: boolean;
  /**
   * `<a:bodyPr><a:prstTxWarp>` (ECMA-376 §20.1.9.19) — WordArt text warp. When
   * present the renderer maps each glyph through the named envelope
   * (presetTextWarpDefinitions) instead of laying text out flat. Omitted from
   * JSON when the body has no warp, so unwarped bodies are byte-identical.
   */
  textWarp?: {
    /** The `prst` name, e.g. `"textArchUp"`, `"textWave1"`. */
    preset: string;
    /** `<a:avLst>` adjust values (adj1, adj2, …) in thousandths of a percent.
     *  Omitted when the author supplied none (preset defaults apply). */
    adj?: number[];
  };
  /**
   * Narrow the inherited `paragraphs` to the PPTX `Paragraph` so consumers see
   * the PPTX-only `eaLnBrk` flag. PPTX `Paragraph extends CoreParagraph`, so
   * this is a covariant refinement of `CoreTextBody.paragraphs`.
   */
  paragraphs: Paragraph[];
}

export interface Presentation {
  slideWidth: number;
  slideHeight: number;
  slides: Slide[];
  /** Theme dk1 color (e.g. "383838"). Used as fallback text color when no explicit color is set. */
  defaultTextColor: string | null;
  /** Theme major (heading) font family name (e.g. "Aptos Display", "Nunito Sans"). Null if not set. */
  majorFont: string | null;
  /** Theme minor (body) font family name (e.g. "Aptos", "Nunito Sans"). Null if not set. */
  minorFont: string | null;
  /** Theme hyperlink colour (hex 6 chars). Used to colour hyperlink runs that have no explicit colour. */
  hlinkColor?: string;
  /** Theme followed-hyperlink colour. Reserved for future visited-link styling. */
  folHlinkColor?: string;
}

export interface Slide {
  index: number;
  /** 1-based slide number (index + 1); used to render slidenum fields */
  slideNumber: number;
  /**
   * The slide's normalized OPC part name (e.g. `ppt/slides/slide3.xml`),
   * resolved through `presentation.xml.rels` in `sldIdLst` order (ECMA-376
   * §19.3.1.42). An internal hyperlink slide jump
   * (`<a:hlinkClick action="ppaction://hlinksldjump" r:id>`, §21.1.2.3.5)
   * carries a rel Target that resolves to this same part name — so
   * {@link PptxPresentation.getSlideIndexByPartName} can turn a click into a
   * slide index. Absent (`undefined`) only for a slide whose part path was not
   * recorded; healthy and broken slides both carry it.
   */
  partName?: string;
  background: Fill | null;
  elements: SlideElement[];
  /** Index-aligned rendered provenance; intentionally carries no editor tree position. */
  elementSources?: SlideElementSource[];
  /**
   * Speaker-notes pane text from `ppt/notesSlides/notesSlideN.xml`
   * (ECMA-376 §13.3.5 — Notes Slide). The full notes-body text as a single
   * string, paragraphs joined with `\n`. Absent (`undefined`) when the slide
   * has no notes part. The renderer ignores this — it is surfaced for tools;
   * read it via {@link PptxPresentation.getNotes}.
   */
  notes?: string;
  /** Classic (`ppt/comments/commentN.xml`, ECMA-376 §13.3.4) or modern
   * threaded slide comments. Omitted when the slide has no comments. */
  comments?: PptxComment[];
  /**
   * `<p:sld show="0">` — the slide is marked hidden in the slide show
   * (ECMA-376 §19.3.1.38). Absent (`undefined`) ⇒ shown. The renderer ignores
   * this; it is a fact surfaced for tools and for {@link PptxViewer}'s hidden-
   * slide modes (read it via `PptxPresentation.isHidden`).
   */
  hidden?: boolean;
  /**
   * RB7 partial degradation: set when this slide's part could not be parsed. The
   * deck still opens with the OTHER slides intact; this one is a placeholder
   * (`elements` empty) whose `parseError` names the offending part (e.g.
   * `"ppt/slides/slide3.xml: <detail>"`). Absent (`undefined`) for every healthy
   * slide. The renderer paints a visible error box instead of slide content.
   */
  parseError?: string;
}

export type SlideElementOrigin = 'master' | 'layout' | 'slide';

export interface SlideElementSource {
  origin: SlideElementOrigin;
}

/**
 * Translucent overlay drawn over a finished slide so it reads faintly
 * (PowerPoint's hidden-slide thumbnail look). A pure render mechanism: the
 * renderer never decides *when* to dim — the caller ({@link PptxViewer}'s
 * `'dim'` mode) does. Both fields are required at the engine boundary; the
 * viewer-facing override (`PptxViewerOptions.hiddenSlideDim`) is partial.
 */
export interface DimOptions {
  /** CSS color of the overlay (e.g. `'#ffffff'`). */
  color: string;
  /** Overlay opacity 0..1 (e.g. `0.6` ⇒ underlying content shows at 40%). */
  opacity: number;
}

/** A slide comment. Supports classic PresentationML comments and modern
 * Microsoft PowerPoint comments (`p188:cm`, [MS-PPTX] §2.16). */
export interface PptxComment {
  /** `<p:cm @authorId>` author-list identifier (ECMA-376 §19.4.1). */
  authorId?: number;
  /** Modern `p188:cm@authorId`. Kept separate so the classic numeric field
   * remains source-compatible. */
  modernAuthorId?: string;
  /** Modern `p188:cm@id`. Classic identity is `(authorId, index)`. */
  id?: string;
  /** `<p:cm @idx>` identifier unique among this author's comments. */
  index?: number;
  /** Resolved author name from `ppt/commentAuthors.xml`. Absent when the
   *  authors file is missing or the `authorId` is out of range. */
  author?: string;
  /** `<p:cm @dt>` — ISO-8601 timestamp the comment was authored. */
  date?: string;
  /** `<p:pos>` anchor on the slide surface, in EMU (ECMA-376 §19.4.5). */
  x?: number;
  y?: number;
  /** Modern comment state. Absent for classic comments. */
  status?: 'active' | 'resolved' | 'closed';
  /** Plain-text comment body (`<p:text>` or modern `<p188:txBody>`). */
  text: string;
  /** Modern replies in authored order. */
  replies?: readonly Readonly<PptxCommentReply>[];
}

export interface PptxCommentReply {
  id?: string;
  authorId?: string;
  author?: string;
  date?: string;
  status?: 'active' | 'resolved' | 'closed';
  text: string;
}

export type SlideElement = ShapeElement | PictureElement | TableElement | ChartElement | MediaElement;

export interface MediaElement {
  type: 'media';
  /** `<p:nvPicPr><p:cNvPr @id>` for the media frame. */
  id?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /** "audio" or "video" */
  mediaKind: 'audio' | 'video';
  /** Poster image zip path (e.g. "ppt/media/image2.png"). Empty when no poster. */
  posterPath: string;
  /** Poster image MIME type (empty when no poster). */
  posterMimeType: string;
  /** Path inside the pptx zip (e.g. "ppt/media/media2.mp4"). Used by getMedia. */
  mediaPath: string;
  /** MIME type of the underlying media (e.g. "audio/mpeg", "video/mp4"). */
  mimeType: string;
}

export interface ShapeElement {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Rotation in degrees, clockwise */
  rotation: number;
  /** Horizontal mirror (a:xfrm flipH) */
  flipH: boolean;
  /** Vertical mirror (a:xfrm flipV) */
  flipV: boolean;
  /** OOXML preset name or "custGeom" when custom paths are used */
  geometry: string;
  fill: Fill | null;
  stroke: Stroke | null;
  textBody: TextBody | null;
  /** `p:cNvSpPr@txBox` — true when this shape is specifically a text box.
   *  A regular shape can still contain text through {@link textBody}. */
  isTextBox?: boolean;
  /** Default text color from p:style > fontRef (hex). Used when run/para has no explicit color. */
  defaultTextColor: string | null;
  /** Custom geometry sub-paths (set only when geometry === "custGeom").
   *  Outer array: one entry per <a:path>; inner: path commands with coords in [0,1]. */
  custGeom: PathCmd[][] | null;
  /** First adjustment value from prstGeom avLst (e.g. trapezoid inset). Range 0–100000. */
  adj: number | null;
  /** Second adjustment value from prstGeom avLst (e.g. arrow head width). Range 0–100000. */
  adj2: number | null;
  /** Third adjustment value from prstGeom avLst (e.g. callout tip x). Range 0–100000. */
  adj3: number | null;
  /** Fourth adjustment value from prstGeom avLst (e.g. callout tip y). Range 0–100000. */
  adj4: number | null;
  /** adj5-adj8: extra polyline vertices for callouts like accentBorderCallout3. */
  adj5: number | null;
  adj6: number | null;
  adj7: number | null;
  adj8: number | null;
  /** Drop shadow from effectLst > outerShdw (null if not present). */
  shadow: Shadow | null;
  /** Inner (inset) shadow from effectLst > innerShdw. ECMA-376 §20.1.8.21. */
  innerShadow?: Shadow;
  /** Coloured glow halo from effectLst > glow. ECMA-376 §20.1.8.17. */
  glow?: Glow;
  /** Soft (feathered) edge — ECMA-376 §20.1.8.31. */
  softEdge?: SoftEdge;
  /** Mirrored reflection — ECMA-376 §20.1.8.27. */
  reflection?: Reflection;
  /** Explicit text frame from a SmartArt drawing's `<dsp:txXfrm>` (absolute EMU,
   *  same space as x/y/width/height). When present the renderer lays text out in
   *  this rectangle instead of the preset/ellipse-derived text rectangle. */
  textRect?: TextRect;
  /** `<a:scene3d>` 3D camera scene (ECMA-376 §20.1.5.5). When the camera is
   *  non-identity the renderer projects the shape through the camera
   *  homography (Phase A). */
  scene3d?: Scene3d;
  /** `<a:sp3d>` 3D shape properties (ECMA-376 §20.1.5.12). Parsed but not
   *  rendered in Phase A. */
  sp3d?: Sp3d;
  /** `<p:nvSpPr><p:cNvPr @id>` — DrawingML cNvPr id. Present for file-authored
   *  shapes (the attribute is schema-required); absent on parser-synthesized
   *  shapes such as the SmartArt data-model fallback. */
  id?: string;
  /** `<p:nvSpPr><p:cNvPr @name>` — author-visible shape name (e.g. "Title 1").
   *  The SmartArt data-model fallback synthesizes `name: "SmartArt"` with no
   *  {@link ShapeElement.id} (see smartart-fallback-contrast.ts). */
  name?: string;
  /** Shape-level hyperlink target resolved from `<p:cNvPr><a:hlinkClick @r:id>`
   *  via slide _rels (ECMA-376 §21.1.2.3.5). For an external link this is the
   *  URL; for an internal slide jump it is the resolved internal part name.
   *  Undefined when the shape carries no hlinkClick. */
  hyperlink?: string;
  /** Raw `<a:hlinkClick @action>` (e.g. `"ppaction://hlinksldjump"`) when the
   *  shape link is an internal PowerPoint action rather than an external URL.
   *  Undefined when absent. */
  hyperlinkAction?: string;
}

/** Absolute text-frame rectangle in EMU (from SmartArt `<dsp:txXfrm>`). */
export interface TextRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

// ===== DrawingML 3D scene (scene3d / sp3d) =====
// 1:1 with the Rust parser's Rot3d / Camera3d / LightRig / Scene3d / Bevel3d /
// Sp3d. Phase A renders only the camera (perspective homography of the planar
// shape); sp3d / lightRig are parsed but rendered in Phase B.

/**
 * 3D rotation in sphere coordinates — ECMA-376 §20.1.5.11 (`CT_SphereCoords`).
 * Angles are in **degrees** (the XML carries 60000ths of a degree; the parser
 * divides once). Per the spec, `lat`/`lon` are latitude/longitude and `rev` is
 * the revolution about the resulting view axis.
 */
export interface Rot3d {
  /** Latitude — rotation about the horizontal (X) axis, degrees. */
  lat: number;
  /** Longitude — rotation about the vertical (Y) axis, degrees. */
  lon: number;
  /** Revolution — in-plane rotation about the view (Z) axis, degrees. */
  rev: number;
}

/**
 * `<a:camera>` — ECMA-376 §20.1.5.5 (`CT_Camera`). `prst` selects one of the
 * 62 preset cameras (§20.1.10.47); `fov`/`zoom`/`rot` optionally override it.
 */
export interface Camera3d {
  /** Preset camera name (`ST_PresetCameraType`), e.g. "perspectiveRelaxed". */
  prst: string;
  /** Field-of-view override in degrees. Omitted = preset default. */
  fov?: number;
  /** Zoom factor as a unit ratio (1.0 = 100%). Omitted = 1.0. */
  zoom?: number;
  /** Camera rotation override. Omitted = preset base orientation. */
  rot?: Rot3d;
}

/**
 * `<a:lightRig>` — ECMA-376 §20.1.5.9 (`CT_LightRig`). Drives the bevel-lip
 * lighting (Phase B): `dir` selects the key-light octant.
 */
export interface LightRig {
  /** Light-rig preset (`ST_LightRigType`), e.g. "threePt". */
  rig: string;
  /** Light direction (`ST_LightRigDirection`): tl/t/tr/l/r/bl/b/br. */
  dir: string;
  /** Optional rotation override of the rig. */
  rot?: Rot3d;
}

/**
 * `<a:scene3d>` — ECMA-376 §20.1.4.1.41 (`CT_Scene3D`). Camera + light rig for
 * a shape's 3D scene.
 */
export interface Scene3d {
  camera: Camera3d;
  lightRig?: LightRig;
}

/**
 * `<a:bevel>` — ECMA-376 §20.1.5.3 (`CT_Bevel`). Lengths in EMU; `w`/`h`
 * default to 76200 EMU and `prst` to "circle".
 */
export interface Bevel3d {
  /** Bevel width in EMU. */
  w: number;
  /** Bevel height in EMU. */
  h: number;
  /** Bevel preset name (`ST_BevelPresetType`). */
  prst: string;
}

/**
 * `<a:sp3d>` — ECMA-376 §20.1.5.12 (`CT_Shape3D`). Rendered in Phase B: bevelT/
 * bevelB are shaded as a lit lip (distance-field + lightRig), extrusionH as a
 * swept side wall, and contour as a flat outline approximation. Numeric fields
 * are omitted from JSON when zero.
 */
export interface Sp3d {
  /** Z position of the front face in EMU (default 0). */
  z?: number;
  /** Extrusion (depth) height in EMU (default 0). */
  extrusionH?: number;
  /** Contour (outline) width in EMU (default 0). */
  contourW?: number;
  /** Contour colour (`<a:contourClr>`, ECMA-376 §20.1.5.12) as a hex string
   *  (e.g. "969696"). Omitted when absent. The renderer draws a flat
   *  approximation of the 3D contour edge (uniform-width outline, no bevel
   *  shading) when both `contourW` and `contourClr` are present. */
  contourClr?: string;
  /** Extrusion side-wall colour (`<a:extrusionClr>`) as a resolved hex string. */
  extrusionClr?: string;
  /** Preset surface material (`ST_PresetMaterialType`), default "warmMatte". */
  prstMaterial: string;
  /** Top bevel. */
  bevelT?: Bevel3d;
  /** Bottom bevel. */
  bevelB?: Bevel3d;
}

export interface TableElement {
  type: 'table';
  /** `<p:nvGraphicFramePr><p:cNvPr @id>` for the table frame. */
  id?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /** Column widths in EMU */
  cols: number[];
  rows: TableRow[];
  /** `<a:tblPr rtl="1">` (ECMA-376 §21.1.3.13): right-to-left table — column 0 at the right edge. */
  rtl?: boolean;
}

export interface TableRow {
  /** Row height in EMU */
  height: number;
  cells: TableCell[];
}

export interface TableCell {
  textBody: TextBody | null;
  fill: Fill | null;
  /** Default run text colour inherited from the table style (`<a:tcTxStyle>`); hex, no `#`. */
  textColor?: string;
  borderL: Stroke | null;
  borderR: Stroke | null;
  borderT: Stroke | null;
  borderB: Stroke | null;
  /** Diagonal from top-left to bottom-right */
  diagonalTL?: Stroke | null;
  /** Diagonal from top-right to bottom-left */
  diagonalTR?: Stroke | null;
  /** Column span */
  gridSpan: number;
  /** Row span */
  rowSpan: number;
  /** Horizontal merge continuation */
  hMerge: boolean;
  /** Vertical merge continuation */
  vMerge: boolean;
}

/**
 * PPTX chart element. The Rust parser emits ChartModel fields flat at the
 * top level, alongside the element position (x/y/width/height in EMU).
 * Pass this straight to `renderChart` from `@silurus/ooxml-core`.
 */
export interface ChartElement {
  type: 'chart';
  /** `<p:nvGraphicFramePr><p:cNvPr @id>` for the chart frame. */
  id?: string;
  /** Frame geometry on the slide, in EMU. */
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /**
   * The chart payload, already in the canonical {@link ChartModel} shape emitted
   * by the Rust parser (`ooxml_common::chart::ChartModel`). Passed straight to
   * `@silurus/ooxml-core`'s `renderChart` — no per-field adapter. The former
   * 60-field flat copy on this interface is gone; all chart properties now live
   * on `chart`.
   */
  chart: ChartModel;
}

export interface PictureElement {
  type: 'picture';
  /** The source drawing node's `<p:cNvPr @id>`. */
  id?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /**
   * Embedded zip path of the raster blip (e.g. "ppt/media/image1.png"). The
   * renderer fetches the bytes lazily by path (see {@link
   * PptxPresentation.getImage}) instead of inlining base64. When the picture is
   * a pure SVG with no raster blip this falls back to the SVG part's path and
   * {@link PictureElement.mimeType} is `image/svg+xml`.
   */
  imagePath: string;
  /** MIME type of the blip at {@link PictureElement.imagePath} (e.g. `image/png`). */
  mimeType: string;
  /**
   * Microsoft 2016 SVG extension (`<a:blip><a:extLst><a:ext
   * uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed>`). When
   * PowerPoint embeds an SVG image, `imagePath` above is only the PNG fallback
   * it rasterizes for compatibility; this is the zip path of the original
   * vector `.svg` part. The renderer prefers this and falls back to the raster
   * if the SVG fails to decode. Omitted when the picture has no svgBlip
   * extension (the common case). Its MIME is always `image/svg+xml` and is
   * owned by the SVG decoder.
   */
  svgImagePath?: string;
  /**
   * Intrinsic pixel width of the raster blip, read from the PNG IHDR at parse
   * time. Omitted for non-PNG payloads. Used internally for the ink-fallback
   * (empty-stroke PNG centering).
   */
  intrinsicWidthPx?: number;
  /** Intrinsic pixel height of the raster blip (PNG IHDR). Omitted for non-PNG. */
  intrinsicHeightPx?: number;
  /**
   * Border line from `<p:pic><p:spPr><a:ln>` (ECMA-376 §20.1.2.2.24). A
   * `p:pic`'s spPr is `CT_ShapeProperties` (§19.3.1.37), so a picture carries
   * the same line model as a shape. `null` when there is no `<a:ln>` or it
   * resolves to `<a:noFill/>` (border explicitly suppressed). The border is
   * stroked along the picture's clip silhouette (roundRect / custGeom / rect).
   */
  stroke: Stroke | null;
  /**
   * `<p:spPr><a:prstGeom prst="…">` preset name (e.g. `"roundRect"`,
   * `"ellipse"`). ECMA-376 §20.1.9.18: a picture's preset geometry is its clip
   * silhouette and the path its border / contour hug. Undefined / omitted = a
   * plain rectangle (`prst="rect"` or no prstGeom). When set, the renderer
   * builds the silhouette via the shared preset-geometry engine (any of the 186
   * presets). `custGeom` takes priority when both are present.
   */
  prstGeom?: string;
  /**
   * Adjust guides from the prstGeom `<a:avLst>` (1/1000-of-a-percent OOXML
   * units), in `gd@name` declaration order (index 0 = adj/adj1, 1 = adj2, …).
   * Omitted when avLst is empty — the preset's own declared defaults then apply.
   */
  prstAdjust?: number[];
  /**
   * ECMA-376 §20.1.8.55 a:srcRect — source image crop as signed fractions of the
   * source width/height, measured from each edge. Omitted when the image
   * is not cropped; when present the parser emits all four edges (absent edges
   * default to 0), so the renderer reads them without a fallback.
   */
  srcRect?: { l: number; t: number; r: number; b: number };
  /** a:blip > a:alphaModFix@amt as 0..1. Undefined = fully opaque. */
  alpha?: number;
  /**
   * ECMA-376 §20.1.8.23 `<a:duotone>` recolour, resolved to its two endpoint
   * colours (through the slide theme). Undefined ⇒ no duotone. When present the
   * renderer decodes the raster once, remaps it along the `clr1`→`clr2`
   * luminance ramp, and caches the recoloured bitmap under a colour-suffixed key.
   */
  duotone?: Duotone;
  /**
   * `<p:spPr><a:custGeom>` clipping path. Same `PathCmd` model as
   * `ShapeElement.custGeom` (one entry per `<a:path>`; coords normalized
   * into [0,1] of the picture's bounding box). The renderer builds a
   * Path2D and `ctx.clip()` before drawing the bitmap so the image is
   * trimmed to the laptop / device silhouette declared in the file.
   */
  custGeom?: PathCmd[][] | null;
  /**
   * Drop shadow from `spPr > effectLst > outerShdw`. A `p:pic`'s `spPr` is
   * `CT_ShapeProperties` (ECMA-376 §19.3.1.37), so the same effects shapes
   * carry apply to images. ECMA-376 §20.1.8.45 (CT_OuterShadowEffect).
   */
  shadow?: Shadow;
  /** Inner (inset) shadow from effectLst > innerShdw. ECMA-376 §20.1.8.40. */
  innerShadow?: Shadow;
  /** Coloured glow halo from effectLst > glow. ECMA-376 §20.1.8.32. */
  glow?: Glow;
  /** Soft (feathered) edge from effectLst > softEdge. ECMA-376 §20.1.8.53. */
  softEdge?: SoftEdge;
  /** Mirrored reflection from effectLst > reflection. ECMA-376 §20.1.8.50. */
  reflection?: Reflection;
  /** `<a:scene3d>` 3D camera scene (ECMA-376 §20.1.5.5). A `p:pic`'s spPr is
   *  `CT_ShapeProperties`, so 3D scenes apply to images. When non-identity the
   *  renderer projects the picture through the camera homography (Phase A). */
  scene3d?: Scene3d;
  /** `<a:sp3d>` 3D shape properties (ECMA-376 §20.1.5.12). Parsed but not
   *  rendered in Phase A. */
  sp3d?: Sp3d;
}
