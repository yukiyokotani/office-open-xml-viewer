// Format-agnostic OOXML types shared across pptx/docx/xlsx packages.
// All positions and sizes are in EMUs (English Metric Units).

import type { MathNode } from './math';
import type { Duotone } from '../image/duotone';
import type { SrcRect } from '../image/crop';

export type PathCmd =
  | { cmd: 'moveTo';     x: number; y: number }
  | { cmd: 'lineTo';     x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'quadBezTo';  x1: number; y1: number; x: number; y: number }
  | { cmd: 'arcTo';      wr: number; hr: number; stAng: number; swAng: number }
  | { cmd: 'close' };

export type Fill = SolidFill | NoFill | GradientFill | PatternFill | ImageFill;

export interface SolidFill {
  fillType: 'solid';
  color: string; // hex 6-char or 8-char (RRGGBBAA with alpha)
}

export interface NoFill {
  fillType: 'none';
}

export interface GradientStop {
  position: number; // 0.0–1.0
  color: string;    // hex 6 or 8 chars
}

export interface GradientFill {
  fillType: 'gradient';
  stops: GradientStop[];
  /** degrees: 0 = left→right, 90 = top→bottom */
  angle: number;
  /** 'linear' | 'radial' */
  gradType: string;
  /** ECMA-376 §20.1.8.41: scale the gradient vector by the fill width/height. */
  scaled?: boolean;
  /** ECMA-376 §20.1.8.46 path shade geometry. */
  path?: 'shape' | 'circle' | 'rect' | string;
  /** Path-gradient focus rectangle, relative to the gradient tile. */
  fillToRect?: FillRect;
  /** Gradient tile rectangle, relative to the shape bounds. */
  tileRect?: FillRect;
  /** Gradient tile mirroring. */
  flip?: 'none' | 'x' | 'y' | 'xy' | string;
  /** Whether the gradient rotates with the containing shape. */
  rotWithShape?: boolean;
}

/**
 * Preset pattern fill — ECMA-376 §20.1.8.40 (CT_PatternFillProperties)
 * with `preset` drawn from §20.1.10.59 (ST_PresetPatternVal).
 */
export interface PatternFill {
  fillType: 'pattern';
  /** Foreground hex colour — used for the "1" pixels of the preset bitmap. */
  fg: string;
  /** Background hex colour — used for the "0" pixels. */
  bg: string;
  /** Preset name, e.g. "pct25", "horz", "diagCross", "lgGrid". */
  preset: string;
}

/**
 * ECMA-376 §20.1.8.30 (CT_RelativeRect) — the destination rectangle a stretched
 * blip is mapped into, as edge insets relative to the fill region. Values are
 * fractions (ST_Percentage / 100000); **negative values let the image bleed
 * past the box (overscan)**. Absent edges default to 0.
 */
export interface FillRect {
  l?: number;
  t?: number;
  r?: number;
  b?: number;
}

/**
 * ECMA-376 §20.1.8.58 (CT_TileInfoProperties) — tiled blip-fill placement.
 * The blip repeats at its native size (scaled by sx/sy) across the fill box.
 * Mutually exclusive with {@link ImageFill.fillRect} (the `stretch` mode).
 */
export interface TileInfo {
  /** Horizontal offset of the first tile, in EMU (`tx`). */
  tx?: number;
  /** Vertical offset of the first tile, in EMU (`ty`). */
  ty?: number;
  /** Horizontal tile scale as a fraction (`sx` / 100000). */
  sx?: number;
  /** Vertical tile scale as a fraction (`sy` / 100000). */
  sy?: number;
  /** Mirror mode: `'none' | 'x' | 'y' | 'xy'` (`flip`, schema default `none`). */
  flip?: string;
  /**
   * Anchor corner the tile grid registers against:
   * `tl|t|tr|l|ctr|r|bl|b|br` (`algn`).
   */
  algn?: string;
}

/**
 * Image fill — ECMA-376 §20.1.8.14 (CT_BlipFillProperties). The embedded blip
 * is carried as a zip path + MIME; the renderer fetches the bytes on demand via
 * {@link RenderOptions.fetchImage} (no base64 inlined at parse time). Both
 * fill-modes are modelled and mutually exclusive: `stretch` (§20.1.8.56) carries
 * {@link ImageFill.fillRect}; `tile` (§20.1.8.58) carries {@link ImageFill.tile}.
 */
export interface ImageFill {
  fillType: 'image';
  /**
   * Embedded zip path of the blip (e.g. "word/media/image1.png"), for the lazy
   * byte-on-demand pipeline. The renderer fetches the bytes via a path-keyed
   * loader ({@link RenderOptions.fetchImage}) instead of inlining base64.
   */
  imagePath: string;
  /** Optional SVG original paired with the raster fallback in `imagePath`. */
  svgImagePath?: string;
  /** MIME type of the blip at {@link ImageFill.imagePath} (e.g. `image/png`). */
  mimeType: string;
  /** `<a:blipFill@dpi>` used to convert native image pixels to a physical tile size. */
  dpi?: number;
  /** `<a:blipFill@rotWithShape>`; absence is preserved because the schema has no default. */
  rotWithShape?: boolean;
  /** `<a:srcRect>` source-image crop, as signed fractions of the blip. */
  srcRect?: SrcRect;
  /**
   * `<a:stretch><a:fillRect>` insets. Absent → fills the whole box (or the
   * fill is tiled — see {@link ImageFill.tile}).
   */
  fillRect?: FillRect;
  /** Presence of `<a:stretch>`; absence is not an implicit stretch default. */
  stretch?: boolean;
  /**
   * `<a:tile>` descriptor. Present only when the blipFill is tiled; mutually
   * exclusive with {@link ImageFill.fillRect}.
   */
  tile?: TileInfo;
  /** `a:blip > a:alphaModFix@amt` as a fraction (0.0–1.0). Absent = opaque. */
  alpha?: number;
  /**
   * ECMA-376 §20.1.8.23 `<a:duotone>` recolour, resolved to its two endpoint
   * colours (through the slide theme). Absent ⇒ no duotone. When present the
   * renderer maps the blip's luminance ramp between the two colours (core
   * `applyDuotone`) — the same recolour a `<p:pic>` duotone applies, wired onto
   * the picture-FILL path (§20.1.8.14) by issue #889.
   */
  duotone?: Duotone;
}

export interface Shadow {
  color: string;  // hex 6 chars
  alpha: number;  // 0.0–1.0
  blur: number;   // EMU
  dist: number;   // EMU
  /** degrees clockwise from East */
  dir: number;
  /** Horizontal scale (1.0 = unchanged). Outer shadows only. */
  sx?: number;
  /** Vertical scale (1.0 = unchanged). Outer shadows only. */
  sy?: number;
  /** Horizontal skew in degrees. Outer shadows only. */
  kx?: number;
  /** Vertical skew in degrees. Outer shadows only. */
  ky?: number;
  /** Alignment origin used before scaling/skewing. Default `b`. */
  algn?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  /** Whether shadow transforms and offset rotate with the shape. Default true. */
  rotWithShape?: boolean;
}

/** ECMA-376 §20.1.8.17 (CT_GlowEffect) — coloured halo with blur radius. */
export interface Glow {
  color: string;  // hex 6 chars
  alpha: number;  // 0.0–1.0
  /** Blur radius in EMU. */
  radius: number;
}

/** ECMA-376 §20.1.8.31 (CT_SoftEdgesEffect) — feather radius in EMU. */
export interface SoftEdge {
  radius: number;
}

/** ECMA-376 §20.1.8.27 (CT_ReflectionEffect) — mirrored copy below the
 *  shape with a linear alpha gradient. Carries the spec attributes whose
 *  defaults the renderer needs to interpret correctly. */
export interface Reflection {
  blur: number;      // EMU
  dist: number;      // EMU
  /** Direction in degrees, clockwise from East. */
  dir: number;
  /** Start alpha (0–1). Default 1.0. */
  stA: number;
  /** Start position along the gradient (0–1). Default 0. */
  stPos: number;
  /** End alpha. Default 0. */
  endA: number;
  /** End position. Default 1.0. */
  endPos: number;
  /** Horizontal scale (1.0 = same width). */
  sx: number;
  /** Vertical scale (-1.0 = full mirror). */
  sy: number;
}

export interface ArrowEnd {
  /** OOXML type: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow" */
  type: string;
  /** Width multiplier: "sm" | "med" | "lg" */
  w: string;
  /** Length multiplier: "sm" | "med" | "lg" */
  len: string;
}

export interface Stroke {
  color: string;
  /** Width in EMU */
  width: number;
  /** Authored non-solid DrawingML line paint. Solid lines use `color`. */
  fill?: Exclude<Fill, { fillType: 'image' } | { fillType: 'none' }>;
  /** OOXML prstDash value: "dash", "dot", "dashDot", "lgDash", "lgDashDot", etc. */
  dashStyle?: string;
  /**
   * DrawingML `<a:custDash>` segments as line-width multipliers
   * (`1` = 100% of the rendered line width). When present this takes
   * precedence over {@link Stroke.dashStyle}.
   */
  customDash?: ReadonlyArray<DrawingMLCustomDashSegment>;
  /** Canvas line cap normalized from DrawingML/VML (`flat` → `butt`). */
  lineCap?: 'butt' | 'round' | 'square';
  /** DrawingML line join (`<a:round>`, `<a:bevel>`, or `<a:miter>`). */
  lineJoin?: 'round' | 'bevel' | 'miter';
  /** Miter limit multiplier (`8` means 800%). Only used with `lineJoin: 'miter'`. */
  miterLimit?: number;
  /** Authored DrawingML pen alignment. Canvas currently paints `ctr` and retains `in`. */
  alignment?: 'ctr' | 'in';
  /** Arrow head at the start of the line */
  headEnd?: ArrowEnd;
  /** Arrow head at the end of the line */
  tailEnd?: ArrowEnd;
  /**
   * ECMA-376 §20.1.8.42 ST_CompoundLine. "sng" (default) | "dbl" |
   * "thinThick" | "thickThin" | "tri". Absent means single line.
   */
  cmpd?: string;
}

/** One `<a:custDash><a:ds>` atom normalized to line-width multipliers. */
export interface DrawingMLCustomDashSegment {
  dash: number;
  space: number;
}

export interface TextBody {
  /** Vertical anchor: "t" | "ctr" | "b" */
  verticalAnchor: string;
  paragraphs: Paragraph[];
  /** Default pt size from lstStyle (overrides renderer default when present) */
  defaultFontSize: number | null;
  /** Inherited bold from layout/master defRPr (null = not set, use false as final default) */
  defaultBold: boolean | null;
  /** Inherited italic from layout/master defRPr (null = not set, use false as final default) */
  defaultItalic: boolean | null;
  /** Text insets in EMU (defaults: lIns=rIns=91440, tIns=bIns=45720) */
  lIns: number;
  rIns: number;
  tIns: number;
  bIns: number;
  /** "square" = wrap, "none" = no wrap */
  wrap: string;
  /** Text direction: "horz" | "vert" | "vert270" | "eaVert" etc. */
  vert: string;
  /** Auto-fit: "sp" = shape grows to fit text, "norm" = font shrinks, "none" = no fit */
  autoFit: string;
  /**
   * `<a:normAutofit fontScale>` (ECMA-376 §21.1.2.1.3) — PowerPoint's stored,
   * pre-computed font-shrink ratio for `autoFit === "norm"`, as a fraction
   * (e.g. 0.625 for `fontScale="62500"`). Null/absent when PowerPoint stored no
   * scale; the renderer then re-derives one. Applying the stored value matches
   * PowerPoint exactly instead of guessing from our own text metrics.
   */
  fontScale?: number | null;
  /** `<a:normAutofit lnSpcReduction>` — stored line-spacing reduction fraction
   *  (e.g. 0.20 for `lnSpcReduction="20000"`). Null/absent when not stored. */
  lnSpcReduction?: number | null;
  /**
   * `<a:bodyPr numCol>` (ECMA-376 §20.1.10.34) — number of text columns inside
   * the shape. Defaults to 1; values > 1 cause the renderer to flow paragraphs
   * across N columns left-to-right, top-to-bottom.
   */
  numCol?: number;
  /** `<a:bodyPr spcCol>` — gap between columns in EMU. Default 0. */
  spcCol?: number;
}

export type SpaceLine =
  | { type: 'pct'; val: number }   // val: e.g. 100000 = 100%, 150000 = 150%
  | { type: 'pts'; val: number };  // val in points

/**
 * A paragraph's bullet marker. For `char`, the marker size is EITHER `sizePct`
 * (a percentage of the run size — ECMA-376 §21.1.2.4.9 `<a:buSzPct>`) OR `sizePts`
 * (an absolute size in points — §21.1.2.4.10 `<a:buSzPts>`), never both: they are
 * the one `EG_TextBulletSize` xsd:choice. `sizePts` is optional (absent when no
 * `<a:buSzPts>` was declared); when present it takes precedence over `sizePct`.
 */
export type Bullet =
  | { type: 'none' }
  | { type: 'inherit' }
  | { type: 'char'; char: string; color: string | null; sizePct: number | null; sizePts?: number; fontFamily: string | null }
  | { type: 'autoNum'; numType: string; startAt: number | null; color: string | null; sizePct?: number | null; sizePts?: number; fontFamily?: string | null };

export interface TabStop {
  /** Position in EMU from the LEADING text-inset edge of the text area —
   *  logical, not physical (ECMA-376 §21.1.2.1): the left edge (after lIns)
   *  in an LTR paragraph, the right edge (before rIns) in an RTL
   *  (`<a:pPr rtl="1">`) paragraph. */
  pos: number;
  /** Alignment: "l" | "r" | "ctr" | "dec" */
  algn: string;
}

export interface Paragraph {
  /** Alignment: "l" | "ctr" | "r" | "just" */
  alignment: string;
  /** Left margin in EMU */
  marL: number;
  /** Right margin in EMU */
  marR: number;
  /** First-line indent in EMU (negative = hanging indent) */
  indent: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  spaceLine: SpaceLine | null;
  /** List nesting level (0–8) */
  lvl: number;
  bullet: Bullet;
  defFontSize: number | null;
  defColor: string | null;
  defBold: boolean | null;
  defItalic: boolean | null;
  defFontFamily: string | null;
  /** Tab stops from pPr > tabLst */
  tabStops: TabStop[];
  /**
   * `<a:pPr rtl="1">` — right-to-left paragraph (ECMA-376 §21.1.2.2.7).
   * When true and no explicit `algn`, the parser-side default flips from
   * "l" to "r"; renderers can also use this flag to flow runs RTL.
   */
  rtl?: boolean;
  runs: TextRun[];
}

export type TextRun = TextRunData | LineBreak | EquationRun;

/**
 * An OMML equation embedded in a paragraph (ECMA-376 §22.1). Parsed into the
 * shared math AST and rendered by `@silurus/ooxml-core`'s math engine.
 * PowerPoint stores these as `a14:m` inside `mc:AlternateContent`.
 */
export interface EquationRun {
  type: 'math';
  /** Parsed OMML node list. */
  nodes: MathNode[];
  /** True for block (`m:oMathPara`) math, false for inline (`m:oMath`). */
  display: boolean;
  /** Paragraph default run size in pt, if declared; absent → renderer inherits. */
  fontSize?: number | null;
  /** Equation colour (hex, no '#') from the math run's rPr; absent → inherit. */
  color?: string | null;
}

export interface TextRunData {
  type: 'text';
  text: string;
  /** null = not set, inherit from paragraph/body defaults */
  bold: boolean | null;
  /** null = not set, inherit from paragraph/body defaults */
  italic: boolean | null;
  underline: boolean;
  /**
   * Specific underline style when not the default single line. Values come
   * from `rPr@u` (ECMA-376 §21.1.2.3.9; ST_TextUnderlineType
   * §20.1.10.82): "dbl", "heavy",
   * "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong",
   * "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash",
   * "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl". Absent means either
   * no underline (when `underline` is false) or the default single line.
   */
  underlineStyle?: string;
  /**
   * Underline-only colour from rPr > uFill (ECMA-376 §21.1.2.3.12). Absent
   * means the underline follows the text colour (uFillTx default).
   */
  underlineColor?: string;
  /** True when rPr strike is sngStrike or dblStrike. */
  strikethrough: boolean;
  /**
   * True only when rPr strike = "dblStrike". Lets the renderer draw two parallel
   * lines instead of one. ECMA-376 §21.1.2.3.9; ST_TextStrikeType
   * §20.1.10.79.
   */
  strikeDouble?: boolean;
  /** Font size in points */
  fontSize: number | null;
  color: string | null;
  fontFamily: string | null;
  /**
   * East Asian font family from rPr > a:ea (ECMA-376 §21.1.2.3.3),
   * resolved through the theme. Renderer uses this for CJK glyphs when
   * present; absent means CJK falls back to fontFamily.
   */
  fontFamilyEa?: string;
  /**
   * Symbol font family from rPr > a:sym (ECMA-376 §21.1.2.3.10), resolved
   * through the theme. PowerPoint stores symbol-font glyphs as Private-Use
   * codepoints U+F020–U+F0FF; the renderer uses this font to resolve them.
   * Absent means no symbol font was declared.
   */
  fontFamilySym?: string;
  /** Baseline shift in thousandths of a point. Positive = superscript, negative = subscript. */
  baseline?: number;
  /**
   * Capitalisation transform from `rPr@cap` — ECMA-376 §21.1.2.3.9;
   * ST_TextCapsType §20.1.10.64.
   * 'all' renders text in upper case; 'small' uses small caps (rendered as
   * upper case at ~80% size when no smcp font feature is available).
   * 'none' or omitted leaves the text unchanged.
   */
  caps?: 'none' | 'small' | 'all';
  /**
   * Inter-character spacing in points. Positive values add space; negative
   * values tighten. DrawingML `rPr@spc` accepts unitless hundredths of a point
   * or an `ST_UniversalMeasure` value (ECMA-376 §21.1.2.3.9; ST_TextPoint
   * §20.1.10.74); the PPTX parser normalizes either form to points.
   */
  letterSpacing?: number;
  /** Set for OOXML field runs (e.g. "slidenum"). When set, renderer replaces text with field value. */
  fieldType?: string;
  /**
   * Hyperlink target resolved from rPr > a:hlinkClick @r:id via the slide's _rels.
   * For an external link this is the URL; for an internal slide jump it is the
   * resolved internal part name (e.g. "../slides/slide3.xml"). Undefined for runs
   * without a hyperlink. ECMA-376 §21.1.2.3.5 (CT_Hyperlink).
   */
  hyperlink?: string;
  /**
   * Raw `<a:hlinkClick @action>` string (e.g. "ppaction://hlinksldjump") when
   * present — its presence marks {@link hyperlink} as an INTERNAL PowerPoint
   * action (slide jump / first / last …) rather than an external URL. Undefined
   * when the hlinkClick has no @action. ECMA-376 §21.1.2.3.5. (IX1)
   */
  hyperlinkAction?: string;
  /**
   * Run-level drop shadow on glyphs (`<a:rPr><a:effectLst><a:outerShdw>`),
   * ECMA-376 §20.1.8.45. Independent of the shape-level shadow on `spPr`.
   * Absent means no run-level shadow.
   */
  shadow?: Shadow;
  /**
   * Mirrored reflection on this run's glyphs
   * (`<a:rPr><a:effectLst><a:reflection>`), ECMA-376 §20.1.8.50.
   * This is independent of a shape-level reflection on `spPr`.
   */
  reflection?: Reflection;
  /**
   * Run-level glyph outline (`<a:rPr><a:ln w="..">`), ECMA-376 §20.1.2.2.24
   * (CT_TextOutlineEffect). Renderer strokes each glyph with the given
   * width / colour in addition to the normal fill. Absent means glyphs are
   * fill-only.
   */
  outline?: TextOutline;
  /**
   * Run-level text highlight / marker colour (`<a:rPr><a:highlight>`),
   * ECMA-376 §21.1.2.3.4. In DrawingML this is a full CT_Color (any
   * srgbClr / schemeClr / sysClr / prstClr + transforms), unlike
   * WordprocessingML's fixed 16-name highlight enum — so the parser already
   * resolves it through the theme/clrMap. The value is a hex string without
   * `#` (6-char opaque, or 8-char RRGGBBAA when an alpha transform applies);
   * the renderer paints a background rectangle behind the run's glyphs.
   * Absent means no highlight.
   */
  highlight?: string;
}

/** Run-level glyph outline. Width is in OOXML EMU (12700 EMU = 1 pt). */
export interface TextOutline {
  width: number;
  /** Hex without '#'. Absent = inherit from text fill colour. */
  color?: string;
}

export interface LineBreak {
  type: 'break';
}

export interface RenderOptions {
  width?: number;
  defaultTextColor?: string | null;
  dpr?: number;
  /** Adaptive decoded-raster memory policy shared by all OOXML renderers. */
  imageResources?: import('../image/adaptive-image-budget.js').ImageResourceOptions;
  majorFont?: string | null;
  minorFont?: string | null;
  /** Theme hyperlink colour (hex 6 chars). Used to colour hyperlink runs without an explicit colour. */
  hlinkColor?: string | null;
  /**
   * Lazily resolve an archive-internal asset (by zip path) to a Blob. The
   * renderer uses this to fetch posters and other large embedded assets on
   * demand, keeping the parse output free of inlined base64.
   */
  fetchMedia?: (path: string) => Promise<Blob>;
  /**
   * Lazily resolve an embedded image (by zip path + MIME) to a Blob. Twin of
   * {@link RenderOptions.fetchMedia} for pictures and blip fills: the renderer
   * fetches raster/SVG bytes on demand and decodes them (`createImageBitmap` /
   * path-keyed `<img>`), so the parse output carries only paths, never base64.
   */
  fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
  /**
   * When true, renderMedia draws only the poster frame — play/pause badges
   * and progress bars are left to the caller. Set by the pptx presentSlide
   * API so its interactive handle can own all control chrome without
   * the static renderer drawing a duplicate play badge.
   */
  skipMediaControls?: boolean;
}
