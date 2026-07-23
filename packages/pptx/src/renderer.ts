import type {
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  MediaElement,
  TableElement,
  ChartElement,
  TableCell,
  Fill,
  TileInfo,
  Stroke,
  TextBody,
  Paragraph,
  TextRun,
  PathCmd,
  Shadow,
  Glow,
  RenderOptions,
  DimOptions,
} from './types';
import { asBullet } from './types';
import {
  renderChart,
  crispOffset,
  buildCustomPath as buildCustomPathCore,
  hexToRgba as hexToRgbaCore,
  resolveFill as resolveFillCore,
  applyStroke as applyStrokeCore,
  buildShapePath,
  EMU_PER_PT as PT_TO_EMU,
  mathToMathML,
  recolorSvg,
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
  renderPresetShape,
  hasPreset,
  buildPresetGeometryPath,
  getConnectorAnchors,
  getCustGeomEndpoints,
  drawArrowHead,
  lineEndRetract,
  retractLineEndpoint,
  computeScene3dQuad,
  isScene3dNonIdentity,
  drawProjected,
  expandProjectedQuad,
  createAuxCanvas,
  applyBevelShading,
  applyExtrusion,
  computeDepthOffset,
  lightDirFromRig,
  materialClass,
  isHTMLCanvas,
  defaultDpr,
  clampCanvasSize,
  classifyCjkFont,
  classifyFontGeneric,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  DEFAULT_KINSOKU_RULES,
  isCjkBreakChar,
  isUax14NoBreakPair,
  containsSeaScript,
  isGraphemeFillText,
  seaMixedBreakOffsets,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
  getCachedSvgImageByPath,
  getCachedBitmapByPath,
  getCachedDuotoneBitmapByPath,
  acquireBitmapCacheLease,
  peekCachedBitmapByPath,
  dropBitmapCacheByPath,
  preferVectorBlip,
  cropSourceRect,
  metafileRasterSize,
  highlightBox,
  symbolFontToUnicode,
  isSymbolFontFamily,
  drawUnderline,
  intendedSingleLinePx,
  rasterHeaderExceedsBudget,
  hasTextWarp,
  buildWarpEnvelope,
  warpGlyphTransform,
  followPathUScale,
} from '@silurus/ooxml-core';
import type { WarpEnvelope, WarpGlyphTransform } from '@silurus/ooxml-core';
import type { CameraInput, Vec2, BevelInput, ExtrusionInput, BevelRegion } from '@silurus/ooxml-core';
import type { MathNode, MathRenderer } from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { classifyPptxHyperlink } from './hyperlink';
import { drawPlayBadge } from './media-chrome';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  type LineVisualOrder,
} from './bidi-line';
import { fitCjkLine, type MeasuredChar } from './cjk-wrap.js';
import { justifyLine, type Justified } from './text-justify';
import { justifiedPiecePositions } from '@silurus/ooxml-core';
import { resolveTableBorderConflict } from './table-border-conflict.js';
import { isSmartArtFallbackShape, smartArtFallbackTextColor } from './smartart-fallback-contrast';
import { resolveTabWidths, type TabItem, type TabStopPx } from './tab-layout.js';
import { drawEaVertRun } from './vertical-text.js';

/** Theme font context threaded through the render call chain. */
export interface RenderContext {
  themeMajorFont: string | null;
  themeMinorFont: string | null;
  /** Theme hyperlink colour as a 6-char hex (no leading #), or null. */
  themeHlinkColor?: string | null;
  /**
   * Device-pixel ratio the slide is being drawn at (the `dpr` passed to
   * `ctx.scale(dpr, dpr)` in the top render fn). Threaded so axis-aligned
   * thin line strokes (table grid borders, text underline / strike-through)
   * can apply {@link crispOffset} and render crisp on a DPR=1 display.
   * Always set — `renderSlide` populates it from the real dpr; the in-module
   * default literals use 1. Required so the crisp-line call sites need no
   * `?? 1` fallback (aligns with docx's required `RenderState.dpr`).
   */
  dpr: number;
  /**
   * Contrast-aware default text colour for THIS slide's synthetic SmartArt
   * fallback shape (issue #805), pre-derived by `renderSlide` from the
   * slide-background luminance via {@link smartArtFallbackTextColor}. Null /
   * absent = keep the ordinary theme default. Consumed only by
   * {@link shapeDefaultTextColor} for shapes matching
   * {@link isSmartArtFallbackShape}; every other shape ignores it.
   */
  smartArtFallbackTextColor?: string | null;
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface PptxTextRunInfo {
  /**
   * Source shape's DrawingML `cNvPr@id`. The identifier is stable and unique
   * within its slide. Absent for parser-synthesized shapes that have no cNvPr.
   */
  shapeId?: string;
  text: string;
  /** X position in CSS px, relative to the shape's top-left corner. */
  inShapeX: number;
  /** Y position (top of line box) in CSS px, relative to the shape's top-left corner. */
  inShapeY: number;
  /** Measured text width in CSS px. */
  w: number;
  /** Line height in CSS px. */
  h: number;
  /** Font size in CSS px. */
  fontSize: number;
  /** CSS `font` shorthand used for canvas drawing (e.g. `"bold 16px Arial"`). */
  font: string;
  /** Shape's left edge in canvas CSS px. */
  shapeX: number;
  /** Shape's top edge in canvas CSS px. */
  shapeY: number;
  /** Shape's width in canvas CSS px. */
  shapeW: number;
  /** Shape's height in canvas CSS px. */
  shapeH: number;
  /** Shape rotation in degrees (clockwise). */
  rotation: number;
  /**
   * Additional rotation from a vertical text body (`vert="vert"` → 90,
   * `vert="vert270"` → -90). The CSS overlay must add this to `rotation`.
   */
  textBodyRotation?: number;
  /**
   * Resolved hyperlink target for this run (IX1), classified into the shared
   * {@link HyperlinkTarget} shape. Present only for runs whose `<a:rPr>` carried
   * an `<a:hlinkClick>`; the overlay makes such spans clickable. The glyph
   * drawing (colour + underline) is unaffected — this is metadata for the
   * transparent overlay only.
   */
  hyperlink?: HyperlinkTarget;
}

export type TextRunCallback = (run: PptxTextRunInfo) => void;

/**
 * Convert EMU to canvas pixels.
 * scale = canvasWidthPx / slideWidthEMU  (so that slideWidth EMU == canvasWidth px)
 */
function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

const hexToRgba = hexToRgbaCore;

/**
 * Paint a run's text-highlight (marker) box behind the glyphs.
 * ECMA-376 §21.1.2.3.4 — `<a:rPr><a:highlight>`. Called before the glyphs are
 * drawn so they (and any underline / strikethrough) sit on top, and before any
 * shadow is set on the context so the box itself isn't shadowed. `width` is the
 * glyph advance computed by the caller (it differs between the normal and
 * tab-stop paths only by the justification stretch added to it). The vertical
 * band comes from the shared `highlightBox` helper. `glyphColor` restores
 * `ctx.fillStyle` so the subsequent fillText draws in the run colour.
 */
export function paintHighlight(
  ctx: CanvasRenderingContext2D,
  x: number,
  baseline: number,
  width: number,
  fontPx: number,
  highlight: string,
  glyphColor: string,
): void {
  const { top, height } = highlightBox(baseline, fontPx);
  ctx.fillStyle = highlight;
  ctx.fillRect(x, top, width, height);
  ctx.fillStyle = glyphColor;
}

/** Simple fill resolver that returns a CSS color string.
 *  For gradient/pattern fills, returns a flat colour (used by table cells etc.,
 *  where we don't have ctx scope to build a CanvasPattern). */
function resolveFill(fill: Fill | null): string | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    return fill.stops.length > 0 ? hexToRgba(fill.stops[0].color) : null;
  }
  // Pattern fills degrade to their foreground colour outside the canvas-aware
  // path; full bitmap rendering happens via resolveShapeFill (core resolveFill).
  if (fill.fillType === 'pattern') return hexToRgba(fill.fg);
  return null;
}

/** Context-aware fill resolver that creates a CanvasGradient for gradient
 * fills and a CanvasPattern for preset pattern fills. */
function resolveShapeFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): string | CanvasGradient | CanvasPattern | null {
  return resolveFillCore(fill, ctx, x, y, w, h);
}

// ===== Text layout helpers =====

// The run-underline painter `drawUnderline` (ECMA-376 §20.1.10.82
// ST_TextUnderlineType) was hoisted to core (`@silurus/ooxml-core`
// `drawUnderline`) so the docx renderer can share the exact same geometry /
// dash dispatch; it is imported at the top of this file and its behaviour is
// byte-identical to the former local copy.

// ── Math (OMML) rendering ──────────────────────────────────────────────────
// Equations are converted to SVG by MathJax once, cached by their MathNode[]
// reference (stable from parse through render), then drawn as images inline.
// Mirrors the docx renderer's pipeline so the typesetting is identical.
interface MathRender {
  /** The equation rasterized as opaque black glyphs on transparent. */
  img: HTMLImageElement;
  /** baseline-relative extents in em (1em = the equation's font size in px). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
  /** Per-colour tinted copies (the black `img` recoloured via source-in). */
  tinted: Map<string, CanvasImageSource>;
}
const mathRenders = new WeakMap<MathNode[], MathRender>();

/**
 * Tint the cached black equation image to `color`. PowerPoint equations follow
 * the run/paragraph text colour, which varies per slide (white on dark, theme
 * accents, …), so we recolour the glyphs at draw time via `source-in` instead
 * of baking a single colour. Cached per colour on the render.
 */
function tintedMathImage(render: MathRender, color: string): CanvasImageSource {
  const cached = render.tinted.get(color);
  if (cached) return cached;
  const iw = render.img.naturalWidth || 1;
  const ih = render.img.naturalHeight || 1;
  const canvas = document.createElement('canvas');
  canvas.width = iw;
  canvas.height = ih;
  const cx = canvas.getContext('2d');
  if (!cx) return render.img;
  cx.drawImage(render.img, 0, 0, iw, ih);
  cx.globalCompositeOperation = 'source-in';
  cx.fillStyle = color;
  cx.fillRect(0, 0, iw, ih);
  render.tinted.set(color, canvas);
  return canvas;
}

function svgToImage(svg: string): Promise<HTMLImageElement> {
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  const img = new Image();
  return new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = url;
  });
}

// Rasterization resolution for equation SVGs, in px per em. A MathJax SVG
// carries its size in `ex` units, so an `<img>` rasterizes it at a small
// intrinsic size and `drawImage` then upscales it — blurry on HiDPI canvases.
// Forcing an explicit px width/height makes the browser rasterize at this
// resolution instead; 256 px/em stays crisp for equations well past 40pt even
// at devicePixelRatio 3 (40pt → ~53px/em × 3 ≈ 160 px/em needed).
const MATH_RASTER_PX_PER_EM = 256;

/** Pin the SVG root to an explicit high-resolution px size before rasterizing. */
function sizeSvgForRaster(svg: string, widthEm: number, heightEm: number): string {
  const w = Math.max(1, Math.round(widthEm * MATH_RASTER_PX_PER_EM));
  const h = Math.max(1, Math.round(heightEm * MATH_RASTER_PX_PER_EM));
  return svg.replace(/<svg([^>]*?)>/, (_m, attrs: string) => {
    const cleaned = attrs.replace(/\s(?:width|height)="[^"]*"/g, '');
    return `<svg${cleaned} width="${w}" height="${h}">`;
  });
}

/** Gather every math run reachable from a slide's shapes and table cells. */
function collectSlideMathRuns(slide: Slide): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  const fromBody = (body: TextBody | null) => {
    if (!body) return;
    for (const para of body.paragraphs) {
      for (const run of para.runs) {
        if (run.type === 'math') found.push({ nodes: run.nodes, display: run.display });
      }
    }
  };
  for (const el of slide.elements) {
    if (el.type === 'shape') fromBody(el.textBody);
    else if (el.type === 'table') {
      for (const row of el.rows) for (const cell of row.cells) fromBody(cell.textBody);
    }
  }
  return found;
}

/**
 * Pre-render every equation on the slide to an image (async), so the
 * synchronous layout/draw passes can place them by reading cached extents.
 * A conversion failure leaves the equation unrendered rather than throwing.
 */
export async function prepareSlideMath(slide: Slide, math: MathRenderer): Promise<void> {
  const runs = collectSlideMathRuns(slide);
  if (runs.length === 0) return;
  await math.loadMathJax();
  for (const r of runs) {
    if (mathRenders.has(r.nodes)) continue;
    try {
      const out = await math.mathMLToSvg(mathToMathML(r.nodes, r.display));
      const sized = sizeSvgForRaster(recolorSvg(out.svg, '#000000'), out.widthEm, out.ascentEm + out.descentEm);
      const img = await svgToImage(sized);
      mathRenders.set(r.nodes, {
        img,
        widthEm: out.widthEm,
        ascentEm: out.ascentEm,
        descentEm: out.descentEm,
        tinted: new Map(),
      });
    } catch {
      // leave unrendered
    }
  }
}

type LayoutSegment = {
  text: string;
  font: string;
  /** Inline DrawingML TAB, classified UAX#9 S during visual ordering (#916). */
  isTab?: true;
  /** Reading-frame gap resolved against a:tabLst immediately before paint. */
  tabWidthPx?: number;
  /**
   * Raw (normalized) font-family requested for this segment's glyphs, kept
   * alongside the composed CSS `font` string so the line-height pass can floor
   * the single-line box to the DOCUMENT font's design line height via core's
   * `intendedSingleLinePx` (ECMA-376 §17.3.1.33 single spacing). Only the
   * tabled substituted faces (Meiryo / Sakkal Majalla) raise the floor; every
   * other family returns 0 and leaves PowerPoint's flat 1.2×em untouched.
   */
  fontFamily?: string;
  sizePx: number;
  color: string;
  underline: boolean;
  /** OOXML rPr @u value when not the default "sng": "dbl"/"dotted"/"wavy"/etc. */
  underlineStyle?: string;
  /** rgba() colour for the underline when uFill overrides the text colour. */
  underlineColor?: string;
  strikethrough: boolean;
  /** Two parallel strike lines (rPr strike="dblStrike"). */
  strikeDouble?: boolean;
  /** Extra inter-character spacing in px (already scaled). */
  letterSpacingPx?: number;
  baseline?: number;
  /** Run-level glyph drop shadow (rPr > effectLst > outerShdw). */
  shadow?: import('@silurus/ooxml-core').Shadow;
  /** Run-level glyph outline (rPr > a:ln). Width in EMU; renderer scales. */
  outline?: import('@silurus/ooxml-core').TextOutline;
  /**
   * Run-level text highlight (rPr > a:highlight, ECMA-376 §21.1.2.3.4) as an
   * rgba() string. Renderer fills a background box behind the glyphs.
   */
  highlight?: string;
  /**
   * Resolved hyperlink target for this segment's run (IX1). Carried so the
   * onTextRun callback can attach it to the overlay span. Does not affect glyph
   * drawing (colour + underline already handled via `color`/`underline`).
   */
  hyperlink?: HyperlinkTarget;
  /**
   * Present when this segment is an OMML equation drawn as an image instead of
   * text (`text` is then ""). Box metrics are in px at the current scale.
   */
  math?: {
    nodes: MathNode[];
    display: boolean;
    width: number;
    ascent: number;
    descent: number;
  };
};

interface LayoutLine {
  segments: LayoutSegment[];
  /** ECMA-376 §21.1.2.2.1 — this line is terminated by a MANUAL line break
   *  (`<a:br>`). In a `just` paragraph it is the end of a logical line and is
   *  left-aligned, not stretched — like the paragraph's last line (§20.1.10.59).
   *  `dist`/`thaiDist` still fill every line, including these. */
  endsWithBreak?: boolean;
}

/**
 * Resolve OOXML theme font references (e.g. "+mn-ea", "+mj-lt") to CSS-safe font names.
 * Canvas will silently ignore an invalid CSS font string, keeping whatever font was set before —
 * leading to wrong text size. Map theme references to generic families as a safe fallback.
 */
function normalizeFontFamily(family: string | null, rc: RenderContext): string {
  if (!family) return rc.themeMinorFont ?? 'sans-serif';
  if (family.startsWith('+')) {
    // +mn-lt = minor Latin, +mj-lt = major Latin, +mn-ea = minor East Asian, +mj-ea = major East Asian
    if (family === '+mj-lt' || family === '+mj-ea' || family === '+mj-cs') {
      return rc.themeMajorFont ?? 'sans-serif';
    }
    // +mn-lt, +mn-ea, +mn-cs, or any other + prefix → minor font
    return rc.themeMinorFont ?? 'sans-serif';
  }
  // OOXML typeface sometimes appends ",<generic>" hint (e.g. "Wingdings,Sans-Serif").
  // Strip it so the CSS font name resolves to the actual named font.
  const primary = family.split(',')[0].trim();
  if (!primary) return rc.themeMinorFont ?? 'sans-serif';
  return primary;
}

/** CSS generic font families — must NOT be quoted in a canvas font string. */
const CSS_GENERIC_FAMILIES = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

/** Infer a CSS generic fallback from a named font so missing fonts degrade
 *  consistently. Delegates the serif/sans/mono decision to the shared core
 *  classifier (single source of truth); maps it to the CSS generic keyword. */
function genericFallback(family: string): string {
  const g = classifyFontGeneric(family);
  return g === 'mono' ? 'monospace' : g === 'serif' ? 'serif' : 'sans-serif';
}

/**
 * Office fonts → metric-compatible, freely-distributable substitutes
 * (`presentation.ts` preloads these webfonts). Putting the substitute in the
 * canvas font stack means a viewer that lacks Calibri/Cambria (macOS, Linux)
 * renders at the SAME advance widths as PowerPoint instead of a wider system
 * serif/sans — without which `wrap="none"` lines overflow the slide. Keys are
 * lower-cased; both the Office face and its substitute share glyph metrics.
 */
const OFFICE_FONT_SUBSTITUTE: Record<string, string> = {
  'calibri': 'Carlito',
  'calibri light': 'Carlito',
  'cambria': 'Caladea',
  'cambria math': 'Caladea',
  // Common Arabic-script faces that hosts rarely ship. Map them to Noto
  // substitutes so RTL slides (e.g. sample-10, which requests Sakkal Majalla /
  // Univers Next Arabic) render with a real web font instead of an oversized
  // OS fallback. "Naskh" covers traditional serif-like Arabic faces; "Sans"
  // covers the modern geometric ones.
  'sakkal majalla': 'Noto Naskh Arabic',
  'traditional arabic': 'Noto Naskh Arabic',
  'simplified arabic': 'Noto Naskh Arabic',
  'arabic typesetting': 'Noto Naskh Arabic',
  'univers next arabic': 'Noto Sans Arabic',
};

/** Generic Arabic fallbacks appended to an Arabic-script font's canvas stack
 *  (before the CSS generic) so Arabic glyphs in an Arabic-targeted family that
 *  the host lacks still resolve to a real Arabic web font when `useGoogleFonts`
 *  is on. */
const ARABIC_FALLBACKS = '"Noto Naskh Arabic", "Noto Sans Arabic"';

/**
 * True when `family` names an Arabic-script face. Only such faces get the Noto
 * Arabic web fonts appended to their canvas stack.
 *
 * These fallback faces (esp. Noto Naskh Arabic) also carry serif-style *Latin*
 * glyphs, so appending them to every font stack made Latin text in an
 * uninstalled Latin/CJK face (e.g. a Japanese gothic carrying "About us") fall
 * into the serif Naskh face instead of degrading to the sans-serif generic.
 * Gating on Arabic-script faces keeps RTL decks (Amiri, Sakkal Majalla, …)
 * correct while letting Latin/CJK text degrade to the right generic.
 */
function isArabicScriptFace(family: string): boolean {
  // Faces we explicitly substitute to a Noto Arabic web font are Arabic-script.
  if (OFFICE_FONT_SUBSTITUTE[family.toLowerCase()]?.includes('Arabic')) return true;
  const l = family.toLowerCase();
  // Common Arabic-script family-name markers (Latin transliterations + Arabic).
  return /arabic|naskh|kufi|nastaliq|amiri|scheherazade|lateef|aldhabi|urdu|farsi|العرب|[؀-ۿ]/.test(l);
}

/** Quote each family for a CSS font-family list (no trailing comma). */
function quoteAll(names: readonly string[]): string {
  return names.map((n) => `"${n}"`).join(', ');
}

/**
 * Build the CSS font-family LIST for a (already-normalized, non-generic) face:
 * named face + metric-compatible Office substitute + script Noto fallbacks +
 * inferred generic. The script fallbacks are:
 *
 * - Arabic-script faces: the Arabic Notos lead (Latin/digits share that face).
 * - CJK faces (Noto KR/SC/TC/JP, ordered by the document's CJK language so
 *   shared Han glyphs take the right shapes — see core/fonts/scripts.ts).
 * - Non-CJK scripts (Cyrillic via Noto Sans/Serif, Thai, Devanagari, Hebrew)
 *   appended unconditionally — no Han collision, browser picks per glyph.
 *
 * Exported for unit testing the fallback ordering.
 */
export function cssFontStack(normalized: string): string {
  const generic = genericFallback(normalized);
  const sub = OFFICE_FONT_SUBSTITUTE[normalized.toLowerCase()];
  const subPart = sub ? `"${sub}", ` : '';
  // Arabic faces keep the historical chain unchanged (Arabic leads; appending a
  // CJK or non-CJK tail would let Latin/digits leak away from the Arabic face).
  if (isArabicScriptFace(normalized)) {
    return `"${normalized}", ${subPart}${ARABIC_FALLBACKS}, ${generic}`;
  }
  const variant: 'sans' | 'serif' = generic === 'serif' ? 'serif' : 'sans';
  const cjk = classifyCjkFont(normalized);
  const cjkPart = cjk ? `${quoteAll(cjkFallbackChain(cjk, variant))}, ` : '';
  const nonCjk = variant === 'serif' ? NON_CJK_SERIF_FALLBACKS : NON_CJK_SANS_FALLBACKS;
  const nonCjkPart = `${quoteAll(nonCjk)}, `;
  return `"${normalized}", ${subPart}${cjkPart}${nonCjkPart}${generic}`;
}

/**
 * Stable string key for a {@link HyperlinkTarget} so adjacent same-format runs
 * that carry the SAME link still merge into one segment (classification mints a
 * fresh object per run, so identity comparison would over-split). `undefined`
 * (no link) and any two links with the same kind + destination compare equal.
 */
function hyperlinkKey(t: HyperlinkTarget | undefined): string {
  if (!t) return '';
  return t.kind === 'external' ? `e:${t.url}` : `i:${t.ref}`;
}

function buildFont(bold: boolean, italic: boolean, sizePx: number, family: string, rc: RenderContext): string {
  const style  = italic ? 'italic ' : '';
  const weight = bold   ? 'bold '   : '';
  const normalized = normalizeFontFamily(family, rc);
  if (CSS_GENERIC_FAMILIES.has(normalized)) {
    return `${style}${weight}${sizePx}px ${normalized}`;
  }
  return `${style}${weight}${sizePx}px ${cssFontStack(normalized)}`;
}

/**
 * Lay out a paragraph into display lines.
 * Handles:
 *  - Explicit line breaks (TextRun type='break')
 *  - Space-based word wrap (Latin text)
 *  - Character-level wrap fallback for CJK / words wider than container
 *  - Tab stops (right-aligned and left-aligned)
 *
 * @param marLPx  Paragraph left margin in canvas px (used for tab stop position calculation)
 */
/** True when the paragraph carries a leading marker (char / autoNum / picture
 *  bullet) that occupies the hanging gutter. A picture (`blip`) bullet lives on
 *  the PPTX-widened {@link Bullet}, so narrow via {@link asBullet} rather than
 *  comparing the raw union member. */
function paragraphHasBullet(para: Paragraph): boolean {
  return (
    para.bullet.type === 'char' ||
    para.bullet.type === 'autoNum' ||
    asBullet(para.bullet).type === 'blip'
  );
}

/** First-line indent (ECMA-376 §21.1.2.2.7 `a:pPr@indent`) resolved to the px
 *  amount the FIRST line's TEXT is shifted right / narrowed by. A positive indent
 *  on a non-bullet paragraph eats into the first line's width and shifts its
 *  start right. Two cases resolve to 0:
 *   - a BULLETED paragraph: `indent` is the marker's hanging gutter, positioned
 *     separately by the bullet/textX geometry (`bulletX = textX + raw indentPx`),
 *     so it does not reduce the text's first-line width here;
 *   - a NEGATIVE indent on a NON-bullet paragraph: clamped to 0. (Known gap:
 *     §21.1.2.2.7 would place such a first line left of marL; honoring that
 *     leftward overhang into the marL gutter is unimplemented. Clamping keeps
 *     the wrap budget and the draw offset in agreement rather than split-brain.)
 *  Shared by the spAutoFit measurement ({@link naturalWidthExceedsBbox}), the
 *  wrap budget ({@link layoutParagraph}) and the draw-side `textXOffset`, so the
 *  three paths can never disagree. */
function firstLineIndentPxFor(hasBullet: boolean, indentPx: number): number {
  return hasBullet ? 0 : Math.max(0, indentPx);
}

/**
 * Decide whether `<a:spAutoFit/>` should let text wrap based on the paragraphs'
 * natural single-line width. Returns true when at least one paragraph would
 * exceed the available text width if rendered without wrapping — that's the
 * cue PowerPoint uses to switch from "grow shape horizontally" to "wrap and
 * grow shape vertically" mode (ECMA-376 §20.1.10.5).
 *
 * The measurement is intentionally conservative: it sums all run widths in a
 * paragraph at the run's own font size, ignoring tab stops and explicit
 * `\t` runs. That matches what PowerPoint compares — "one continuous line
 * of glyphs" — and avoids over-eager wrap when a paragraph is barely wider
 * than the bbox due to font-measurement drift.
 *
 * Known limitation (issue #1006): because tab jumps are ignored here, a tab-led
 * paragraph inside a `spAutoFit` shape whose raw glyph sum fits the box will
 * take the grow-not-wrap path even if its tab-jumped extent would overflow. The
 * adjudicated fixtures use `noAutofit`; spAutoFit + tab is unadjudicated and left
 * as documented behaviour rather than guessed.
 */
export function naturalWidthExceedsBbox(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bw: number,
  lPad: number,
  rPad: number,
  scale: number,
  rc: RenderContext,
): boolean {
  const bodyDefaultFontPx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;
  for (const para of body.paragraphs) {
    const marLPx = emuToPx(para.marL, scale);
    const marRPx = emuToPx(para.marR, scale);
    const indentPx = emuToPx(para.indent, scale);
    // The first-line indent is consumed only by a NON-bullet paragraph and only
    // when positive (firstLineIndentPxFor) — the SAME amount the wrap and draw
    // passes use, so the measurement can't disagree with what actually renders.
    const firstLineIndent = firstLineIndentPxFor(paragraphHasBullet(para), indentPx);
    const textMaxW = bw - lPad - rPad - marLPx - marRPx - firstLineIndent;
    let lineW = 0;
    for (const run of para.runs) {
      if (run.type !== 'text') continue;
      const sizePx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale
        : (para.defFontSize != null
            ? para.defFontSize * PT_TO_EMU * scale
            : bodyDefaultFontPx);
      const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
      const isBold = run.bold ?? para.defBold ?? body.defaultBold ?? false;
      const isItalic = run.italic ?? para.defItalic ?? body.defaultItalic ?? false;
      ctx.font = buildFont(isBold, isItalic, sizePx, family, rc);
      lineW += ctx.measureText(run.text).width;
      if (lineW > textMaxW) return true;
    }
  }
  return false;
}

/** True when any character of `s` is a CJK / ideographic glyph that allows a
 *  per-character line break (see core's `isCjkBreakChar` for the ranges). */
function tokenHasCjk(s: string): boolean {
  for (const ch of s) {
    if (isCjkBreakChar(ch.codePointAt(0) ?? 0)) return true;
  }
  return false;
}

/** Number of Unicode code points in `s` (NOT UTF-16 code units). OOXML letter
 *  spacing (rPr @spc) adds advance per GLYPH, and a supplementary-plane CJK
 *  ideograph (e.g. 𠮟 U+20B9F) is one glyph encoded as a surrogate pair. The
 *  per-glyph draw loops (`for (const ch of text)`) and core's
 *  `justifiedPiecePositions` both iterate code points, so the letter-spacing
 *  width math must count code points too — using `text.length` (code units)
 *  would over-count by one per surrogate pair and drift the pen. */
function codePointCount(s: string): number {
  let n = 0;
  for (const _ of s) n++;
  return n;
}

export function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  para: Paragraph,
  maxWidthPx: number,
  defaultFontSizePx: number,
  defaultColor: string,
  scale: number,
  marLPx: number,
  defaultBold: boolean = false,
  defaultItalic: boolean = false,
  fontScale: number = 1.0,
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  firstLineIndentPx: number = 0,
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  // The first line's wrap budget is narrower by a POSITIVE first-line indent
  // (it occupies indentPx of the line); continuation lines use the full width.
  // `lines.length === 0` ⇒ still filling the first line (newLine() pushes to it).
  const lineMaxW = () => maxWidthPx - (lines.length === 0 ? firstLineIndentPx : 0);
  let currentLine: LayoutLine = { segments: [] };
  let lineW = 0; // current line's accumulated width
  // ECMA-376 §17.18.93 ST_TextWrappingType "square" is whitespace-aware: a
  // non-whitespace token is never broken away from the preceding non-whitespace
  // content. We only allow a wrap before a token if at least one whitespace
  // run has appeared on the current line — otherwise the line overflows the
  // shape (PowerPoint's actual behavior, e.g. "YoY+11.9%" mixed-size runs in
  // sample-2 slide-7 stay on one line even though the bbox is tight).
  let hasWhitespaceOnLine = false;

  // ── Wrap-aware tab context (issue #1006) ──────────────────────────────────
  // A tab is a horizontal pen JUMP to its stop within the visual line where it
  // occurs; content overflowing the line's right edge wraps normally and every
  // CONTINUATION line re-anchors at the leading text-inset edge (text-left).
  // The wrap budget must ACCOUNT for the tab jump, so a tabbed line measures its
  // NATURAL extent with the SAME resolver the paint pass uses (`resolveTabWidths`
  // with an infinite limit — the #835 clamp must not hide overflow). Tab-free
  // lines keep the fast additive `lineW` path unchanged (byte-identical).
  const baseRtl = para.rtl === true;
  const marRPxL = emuToPx(para.marR, scale);
  const stopsPxL: TabStopPx[] = (para.tabStops ?? []).map((s) => ({
    pos: emuToPx(s.pos, scale),
    algn: s.algn,
  }));
  // Effective default tab grid (§21.1.2.2.7): explicit pPr value, else the
  // PowerPoint universal 1-inch default so a `\t` never collapses to a space.
  const defTabSzPxL = emuToPx(para.defTabSz ?? 914400, scale);
  // A tab contributes nothing to the additive `lineW` (its gap is resolved), so
  // track "the line carries a tab" explicitly rather than via lineW.
  let lineHasTab = false;
  // Logical-order items for the current line (content advances + tab markers),
  // mirroring the paint-side items so layout and paint resolve identically.
  let lineItems: TabItem[] = [];
  // Space width for the (now unused when defTabSz>0) no-stop fallback; captured
  // from the first tab's font, matching the paint pass.
  let tabSpaceW = 0;

  /** Leading pen for the current line in the reading frame, matching the paint
   *  pass's `leadingIndentPx`: RTL right-anchors at marR (no first-line indent);
   *  LTR starts at marL plus the first line's positive indent. */
  const lineStartPen = (): number =>
    baseRtl ? marRPxL : marLPx + (lines.length === 0 ? firstLineIndentPx : 0);

  /** Natural extent (px advance from the line's leading pen) of the current
   *  line's committed items plus an optional trailing content advance, resolved
   *  against the tab grid with NO trailing clamp. */
  const tabAwareExtent = (extraW = 0): number => {
    const items = extraW > 0 ? [...lineItems, { isTab: false, width: extraW }] : lineItems;
    const widths = resolveTabWidths(items, stopsPxL, lineStartPen(), Infinity, tabSpaceW, defTabSzPxL);
    let sum = 0;
    for (const w of widths) sum += w;
    return sum;
  };

  /** Whether a content advance of `w` still fits the current line's budget.
   *  Tab-free lines use the fast additive path (byte-identical); a tabbed line
   *  resolves the WHOLE hypothetical line (candidate-aware) so a right/centre
   *  tab — whose gap SHRINKS as its cell grows — is placed correctly rather than
   *  via `lineW + w`. An infinite budget (wrap="none" / spAutoFit measured on
   *  one line) never wraps, so short-circuit before the O(items) resolve. */
  const fitsW = (w: number): boolean => {
    const budget = lineMaxW();
    if (!Number.isFinite(budget)) return true;
    return lineHasTab ? tabAwareExtent(w) <= budget : lineW + w <= budget;
  };

  /** Max additional content advance the current line can accept before its
   *  natural extent exceeds the budget. Tab-free ⇒ the additive remainder; a
   *  tabbed line ⇒ the exact monotone threshold of `tabAwareExtent` (correct for
   *  left / right / centre tabs, so a fitting right-tab cell is never wrapped
   *  early). Used by the CJK / SEA prefix-fit, which take a scalar budget. */
  const availW = (): number => {
    const budget = lineMaxW();
    if (!lineHasTab) return budget - lineW;
    if (!Number.isFinite(budget)) return Infinity;
    if (tabAwareExtent(0) >= budget) return 0;
    let lo = 0;
    let hi = budget;
    for (let i = 0; i < 40; i++) {
      const mid = (lo + hi) / 2;
      if (tabAwareExtent(mid) <= budget) lo = mid;
      else hi = mid;
    }
    return lo;
  };

  const newLine = (endsWithBreak = false) => {
    if (endsWithBreak) currentLine.endsWithBreak = true;
    lines.push(currentLine);
    currentLine = { segments: [] };
    lineW = 0;
    lineHasTab = false;
    lineItems = [];
    hasWhitespaceOnLine = false;
  };

  const push = (
    text: string,
    font: string,
    sizePx: number,
    color: string,
    underline: boolean,
    strikethrough: boolean,
    baseline?: number,
    extras?: {
      strikeDouble?: boolean;
      letterSpacingPx?: number;
      underlineStyle?: string;
      underlineColor?: string;
      shadow?: import('@silurus/ooxml-core').Shadow;
      outline?: import('@silurus/ooxml-core').TextOutline;
      highlight?: string;
      /** Raw normalized family for the design-line-height floor (see LayoutSegment). */
      fontFamily?: string;
      /** Resolved hyperlink target (IX1) — passed through to the overlay span. */
      hyperlink?: HyperlinkTarget;
    },
  ) => {
    if (!text) return;
    ctx.font = font;
    const lsPx = extras?.letterSpacingPx ?? 0;
    const baseW = ctx.measureText(text).width;
    // Letter spacing adds an extra gap between every character, including
    // after the last one — matches the "advance width" semantics of OOXML
    // spc (each glyph's advance grows by spc points). Tab stops measure the
    // same way so this stays consistent with measureText below.
    const w = baseW + lsPx * codePointCount(text);
    const strikeDouble = extras?.strikeDouble;
    const underlineStyle = extras?.underlineStyle;
    const underlineColor = extras?.underlineColor;
    const shadow = extras?.shadow;
    const outline = extras?.outline;
    const highlight = extras?.highlight;
    const fontFamily = extras?.fontFamily;
    const hyperlink = extras?.hyperlink;
    // Shadow / outline use object identity for merging — adjacent runs share
    // the same object since the run is parsed once. Different objects (or
    // one set / one missing) force a new segment.
    const sameMeta = (a: LayoutSegment) =>
      !a.math &&
      !a.isTab &&
      a.font === font &&
      a.color === color &&
      a.underline === underline &&
      (a.underlineStyle ?? '') === (underlineStyle ?? '') &&
      (a.underlineColor ?? '') === (underlineColor ?? '') &&
      a.strikethrough === strikethrough &&
      (a.strikeDouble ?? false) === (strikeDouble ?? false) &&
      (a.letterSpacingPx ?? 0) === lsPx &&
      a.baseline === baseline &&
      a.shadow === shadow &&
      a.outline === outline &&
      (a.highlight ?? '') === (highlight ?? '') &&
      (a.fontFamily ?? '') === (fontFamily ?? '') &&
      hyperlinkKey(a.hyperlink) === hyperlinkKey(hyperlink);
    lineW += w;
    // Mirror the paint-side item sequence so a tabbed line's wrap budget resolves
    // identically (issue #1006). One item per push call; consecutive content
    // items simply sum inside resolveTabWidths.
    lineItems.push({ isTab: false, width: w });
    const last = currentLine.segments.at(-1);
    if (last && sameMeta(last)) {
      last.text += text;
    } else {
      currentLine.segments.push({ text, font, fontFamily, sizePx, color, underline, underlineStyle, underlineColor, strikethrough, strikeDouble, letterSpacingPx: lsPx || undefined, baseline, shadow, outline, highlight, hyperlink });
    }
  };

  // UAX#14 LB13 (行頭禁則): pull the trailing word of the current line down onto a
  // fresh line so a glued non-starter (comma, period, … in a SEPARATE run, no
  // whitespace between) does not orphan at the next line's head nor tear the word.
  // Re-pushes the word — with its run formatting — to lead the new line; the
  // caller then appends the non-starter. The trailing word normally lives in the
  // last (same-meta-merged) segment and is split at its last whitespace; when a
  // formatting change split the word across segments, the last segment has no
  // internal whitespace, so the whole tail segment moves down instead (the word
  // splits at the format seam, but the comma is still never orphaned — matching
  // docx/xlsx). Returns false (changing nothing) only when that tail segment IS
  // the whole line (no preceding content / nowhere to retract to). Mirrors the
  // docx/xlsx fixes; the ASCII non-starters live in
  // DEFAULT_KINSOKU_RULES.lineStartForbidden.
  const retractTrailingWord = (): boolean => {
    const seg = currentLine.segments.at(-1);
    if (!seg || seg.math) return false;
    const m = /^(.*\s)(\S+)$/s.exec(seg.text);
    let word: string;
    if (m) {
      seg.text = m[1]; // close the current line on the whitespace boundary
      word = m[2];
    } else if (currentLine.segments.length > 1) {
      currentLine.segments.pop(); // tail segment of a format-split word moves whole
      word = seg.text;
    } else {
      return false; // the segment is the whole line — cannot retract without emptying it
    }
    newLine();
    // Re-push the word (with its run formatting) so it leads the fresh line.
    push(word, seg.font, seg.sizePx, seg.color, seg.underline, seg.strikethrough, seg.baseline, {
      strikeDouble: seg.strikeDouble,
      letterSpacingPx: seg.letterSpacingPx,
      underlineStyle: seg.underlineStyle,
      underlineColor: seg.underlineColor,
      shadow: seg.shadow,
      outline: seg.outline,
      highlight: seg.highlight,
      fontFamily: seg.fontFamily,
    });
    return true;
  };

  for (const run of para.runs) {
    if (run.type === 'break') {
      // The line being closed ends at a MANUAL break (§21.1.2.2.1) — mark it so
      // a `just` paragraph left-aligns it like its last line (§20.1.10.59).
      newLine(true);
      continue;
    }

    // ── OMML equation ─────────────────────────────────────────────────────
    if (run.type === 'math') {
      const render = mathRenders.get(run.nodes);
      // Equation font size: explicit run size (pt→px) else paragraph default.
      const emPx = run.fontSize != null
        ? run.fontSize * PT_TO_EMU * scale * fontScale
        : defaultFontSizePx;
      const width = render ? render.widthEm * emPx : 0;
      const ascent = render ? render.ascentEm * emPx : 0;
      const descent = render ? render.descentEm * emPx : 0;
      // Block (display) math gets its own line; the draw pass centres it.
      if (run.display && lineW > 0) newLine();
      else if (!fitsW(width) && lineW > 0) newLine();
      lineItems.push({ isTab: false, width });
      currentLine.segments.push({
        text: '',
        font: `${emPx}px sans-serif`,
        sizePx: emPx,
        // Equations follow their own run colour (e.g. a purple title); the
        // draw pass tints the glyph image to this colour. Fall back to the
        // paragraph/body default when the run carries no explicit colour.
        color: run.color ? hexToRgba(run.color) : defaultColor,
        underline: false,
        strikethrough: false,
        math: { nodes: run.nodes, display: run.display, width, ascent, descent },
      });
      lineW += width;
      if (run.display) newLine();
      continue;
    }

    const sizePx = run.fontSize != null ? run.fontSize * PT_TO_EMU * scale * fontScale : defaultFontSizePx;
    // Font family cascade: run → paragraph defFontFamily → theme minor font → 'sans-serif'
    const family = normalizeFontFamily(run.fontFamily ?? para.defFontFamily ?? null, rc);
    // East Asian font (rPr > ea) — used for CJK glyphs when set; otherwise
    // CJK characters reuse the latin font. ECMA-376 §21.1.2.3.7.
    const familyEa = run.fontFamilyEa
      ? normalizeFontFamily(run.fontFamilyEa, rc)
      : null;
    // Symbol font (rPr > a:sym) — used for Private-Use symbol glyphs (U+F0xx).
    const familySym = run.fontFamilySym
      ? normalizeFontFamily(run.fontFamilySym, rc)
      : null;
    // Hyperlink runs without an explicit colour pick up the theme hlink colour
    // (ECMA-376 §20.1.2.3.5 — hyperlinks inherit theme hyperlink slot).
    let color: string;
    if (run.color) {
      color = hexToRgba(run.color);
    } else if (run.hyperlink && rc.themeHlinkColor) {
      color = hexToRgba(rc.themeHlinkColor);
    } else {
      color = defaultColor;
    }
    // Cascade: run → paragraph defRPr → body/layout default → false
    const isBold   = run.bold   ?? para.defBold   ?? defaultBold;
    const isItalic = run.italic ?? para.defItalic ?? defaultItalic;
    const font   = buildFont(isBold, isItalic, sizePx, family, rc);
    const fontEa = familyEa
      ? buildFont(isBold, isItalic, sizePx, familyEa, rc)
      : font;
    ctx.font = font;

    // ECMA-376 §21.1.2.3.13 — caps transforms the rendered glyphs without
    // changing the underlying text. "small" emulated as upper-case glyphs at
    // ~80% size is the long-established Office fallback when the font lacks
    // smcp; we just upper-case for now and rely on the configured size.
    const caps = run.caps;
    let baseText = run.text;
    if (caps === 'all' || caps === 'small') baseText = baseText.toUpperCase();

    // Resolve field values (e.g. slidenum → actual slide number)
    const runText = (run.fieldType === 'slidenum' && slideNumber !== undefined)
      ? String(slideNumber)
      : baseText;

    // Hyperlink runs render underlined unless an explicit u attribute already
    // says otherwise. Spec: ECMA-376 §20.1.2.3.5 (hyperlinks default to the
    // hlink character style, which underlines).
    const segUnderline = run.underline || (run.hyperlink !== undefined);
    const segStrikeDouble = run.strikeDouble === true;
    // letterSpacing arrives in points; convert to canvas px using the same
    // EMU→px scale the renderer applies to font sizes.
    const lsPx = run.letterSpacing != null ? run.letterSpacing * PT_TO_EMU * scale : 0;
    const segExtras = {
      strikeDouble: segStrikeDouble,
      letterSpacingPx: lsPx,
      underlineStyle: run.underlineStyle,
      underlineColor: run.underlineColor ? hexToRgba(run.underlineColor) : undefined,
      shadow: run.shadow,
      outline: run.outline,
      // Raw latin/primary family for the design-line-height floor. CJK per-char
      // pushes below override this to `familyEa` when they draw with `fontEa`,
      // so a Meiryo set only as the East Asian typeface is still floored.
      fontFamily: family,
      // §21.1.2.3.4 — highlight is a resolved hex (6-char opaque or 8-char
      // RRGGBBAA); hexToRgba handles both, matching how text/underline colours
      // are converted for canvas.
      highlight: run.highlight ? hexToRgba(run.highlight) : undefined,
      // IX1 — classify the resolved hyperlink target string into the shared
      // HyperlinkTarget shape (external URL vs internal slide jump). The core
      // TextRun type carries only `hyperlink` (no action field), so the string
      // alone drives classification: a ppaction://… or a scheme-less internal
      // part name is treated as internal. Overlay-only; does not affect glyphs.
      hyperlink: classifyPptxHyperlink(run.hyperlink),
    };

    // Split on whitespace boundaries, keeping the whitespace tokens
    const tokens = runText.split(/(\s+)/);

    for (const token of tokens) {
      if (!token) continue;

      // ── Tab character ────────────────────────────────────────────────────
      if (/^\t+$/.test(token)) {
        // §21.1.2.1.x: retain every tab inline. Gap resolution is deferred until
        // paint, when every cell width is known; UAX#9 S then reorders cells.
        // The wrap pass measures the tab jump via `tabAwareExtent` (#1006) so the
        // line breaks at the correct point, and a continuation line (which never
        // carries the tab) re-anchors at text-left.
        if (!lineHasTab) {
          ctx.font = font;
          tabSpaceW = ctx.measureText(' ').width;
        }
        for (const _ of token) {
          currentLine.segments.push({
            text: '',
            isTab: true,
            font,
            fontFamily: family,
            sizePx,
            color,
            underline: false,
            strikethrough: false,
          });
          lineItems.push({ isTab: true, width: 0 });
        }
        lineHasTab = true;
        continue;
      }

      ctx.font = font;
      const tokW = ctx.measureText(token).width;
      const isWhitespace = /^\s+$/.test(token);

      // ── Symbol-font characters (Wingdings/Webdings/Symbol) ───────────────
      // PowerPoint stores symbol glyphs as Private-Use codepoints U+F020–U+F0FF
      // and picks the font via rPr > a:sym (ECMA-376 §21.1.2.3.10). Map the
      // known ones to Unicode equivalents so they render reliably regardless of
      // whether the symbol font is installed; fall back to the real symbol font
      // for unmapped glyphs.
      const SYMBOL_PUA_RE = /[-]/;
      // Gate on core's isSymbolFontFamily (exact "symbol" / any "wingdings";
      // shared with docx). A familySym (a:sym, §21.1.2.3.10) explicitly names
      // the run's symbol typeface, so its presence also opens the path.
      // (Webdings / "SymbolMT" no longer match the family branch — both already
      // passthrough unchanged since core gates "symbol" exactly and has no
      // Webdings table, so this is behaviour-preserving.)
      if (SYMBOL_PUA_RE.test(token) && (familySym != null || isSymbolFontFamily(family))) {
        const symName = familySym ?? family;
        for (const ch of token) {
          let drawCh = ch;
          let chFont = font;
          if (SYMBOL_PUA_RE.test(ch)) {
            const mapped = symbolFontToUnicode(ch, symName);
            if (mapped !== ch) {
              drawCh = mapped;
              chFont = buildFont(isBold, isItalic, sizePx, 'sans-serif', rc);
            } else {
              chFont = buildFont(isBold, isItalic, sizePx, symName, rc);
            }
          }
          ctx.font = chFont;
          const chW = ctx.measureText(drawCh).width;
          if (!fitsW(chW) && lineW > 0) newLine();
          push(drawCh, chFont, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        }
        continue;
      }

      // CJK characters allow line-breaking at any character boundary (no whitespace
      // needed). When a token contains CJK, wrap character-by-character so that CJK
      // text flows onto the same line as preceding Latin text (e.g. "EC市場で…").
      // Per-character font dispatch picks `fontEa` for CJK glyphs when the run
      // declared an explicit East Asian typeface (rPr > ea); other characters
      // keep the Latin font so the latin/ea boundary mid-token stays clean.
      const hasCJK = tokenHasCjk(token);
      // Issue #960 — a token that mixes CJK with SEA (Thai/Lao/Khmer) must NOT
      // take the CJK-only path (which tears the SEA interior at arbitrary
      // character boundaries): route it to the unified SEA branch below, whose
      // offset set merges the CJK per-character opportunities with the SEA
      // dictionary/transition ones so each script keeps its own break rule. The
      // exception is `eaLnBrk=false` (§21.1.2.2.7), which forbids breaking the
      // East Asian word at all — keep that on the CJK path so the token stays
      // whole. A CJK token with no SEA is unchanged.
      const routeCjk = hasCJK && (!containsSeaScript(token) || para.eaLnBrk === false);
      if (routeCjk) {
        // Measure each grapheme with its per-char font (latin/ea boundary stays
        // clean), then place chars according to a:pPr@eaLnBrk (ECMA-376
        // §21.1.2.2.7, "East Asian Line Break"):
        //   • eaLnBrk=true (default) → East Asian text MAY break at character
        //     boundaries, so we wrap char-by-char with kinsoku (§17.15.1.58–.60):
        //     forbidden leaders never start a line and forbidden followers never
        //     end one. fitCjkLine reuses core's kinsokuAdjustedSplit.
        //   • eaLnBrk=false → an East Asian word must NOT be split mid-character.
        //     The whole token moves to a fresh line if it doesn't fit, but is
        //     never torn; when wider than the line it overflows and the shape's
        //     existing clipping handles it.
        //
        // DEFAULT_KINSOKU_RULES is correct for pptx: PresentationML has no custom
        // forbidden-set element (w:noLineBreaksBefore/After are WordprocessingML-only).
        // docx's analogous CJK path (renderer.ts, fitCJKPrefix) is intentionally
        // separate: substring binary-search fit + cross-run 追い出し. Do not unify them.
        const measured: (MeasuredChar & { font: string; family: string })[] = [];
        for (const ch of token) {
          const isEa = isCjkBreakChar(ch.codePointAt(0) ?? 0) && familyEa != null;
          const chFont = isEa ? fontEa : font;
          // Floor to the family actually rendering this glyph: `familyEa` for
          // CJK when an East Asian typeface was declared, else the latin family.
          const chFamily = isEa ? (familyEa as string) : family;
          ctx.font = chFont;
          measured.push({ ch, w: ctx.measureText(ch).width, font: chFont, family: chFamily });
        }
        if (para.eaLnBrk === false) {
          // Keep the East Asian word whole. If the current line already has
          // content and the token would overflow, wrap once before placing it;
          // never break mid-token (an over-wide token simply overflows).
          const tokenW = measured.reduce((acc, m) => acc + m.w, 0);
          if (lineW > 0 && !fitsW(tokenW)) newLine();
          for (const m of measured) {
            push(m.ch, m.font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, { ...segExtras, fontFamily: m.family });
          }
          continue;
        }
        let rest = measured;
        while (rest.length > 0) {
          // Effective start pen so the fit sees the line's true remaining width:
          // budget − availW = the additive pen (tab-free / left tab) or the
          // slack-adjusted pen (right / centre tab). Equivalent to lineW when no
          // tab, so tab-free CJK is byte-identical.
          const cjkPen = Number.isFinite(lineMaxW()) ? lineMaxW() - availW() : lineW;
          let n = fitCjkLine(rest, cjkPen, lineMaxW(), DEFAULT_KINSOKU_RULES);
          if (n === 0) {
            // A non-empty line can't take the run head → break and retry empty.
            // But a line holding ONLY a tab (lineW===0, tab jump consumed the
            // width) must NOT be finalised as a tab-only line — place one glyph
            // so it overflows, mirroring the Latin "no break opportunity" rule.
            if (lineW > 0) { newLine(); continue; }
            n = 1;
          }
          for (let i = 0; i < n; i++) {
            const m = rest[i];
            push(m.ch, m.font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, { ...segExtras, fontFamily: m.family });
          }
          rest = rest.slice(n);
          if (rest.length > 0) newLine();
        }
        continue;
      }

      // SEA (Thai/Lao/Khmer) line breaking (issue #797 / #960). These scripts have
      // no inter-word spaces, so `token` is a whole run; break it only at a member
      // of `seaBreaks` — the UNION of dictionary word boundaries, the no-space
      // SEA↔non-SEA script transitions, and (for a mixed CJK+SEA token routed here
      // from above) the CJK per-character opportunities, kinsoku-filtered. A SEA
      // token with no boundary (single over-long word / Segmenter unavailable)
      // still routes here so its emergency split stays grapheme-safe.
      if (containsSeaScript(token)) {
        const seaBreaks = seaMixedBreakOffsets(token, { cjk: true, kinsoku: DEFAULT_KINSOKU_RULES });
        // A line-piece may mix a Thai run (Latin/Thai `font`) with a CJK run
        // (`fontEa` when an East Asian typeface is declared). Measure each maximal
        // same-font sub-run WHOLE — per-char measurement would break Thai shaping —
        // and push each with its own font so CJK glyphs get `fontEa`. For a
        // pure-SEA piece (or `familyEa` absent / resolving to the SAME font) this
        // is one run with `font`: whole-string measure + single push, i.e.
        // byte-identical to the pre-#960 path. The split is taken ONLY when the
        // EA font actually differs — otherwise measuring per-run and re-merging
        // identical-font pushes could disagree on cross-boundary shaping.
        const eaDiffers = familyEa != null && fontEa !== font;
        const isEaCh = (ch: string): boolean => eaDiffers && isCjkBreakChar(ch.codePointAt(0) ?? 0);
        const measureSub = (sub: string): number => {
          let w = lsPx * codePointCount(sub);
          let runText = '';
          let runEa: boolean | null = null;
          const flush = (): void => {
            if (runText === '') return;
            ctx.font = runEa ? (fontEa as string) : font;
            w += ctx.measureText(runText).width;
            runText = '';
          };
          for (const ch of sub) {
            const ea = isEaCh(ch);
            if (runEa === null || ea === runEa) { runText += ch; runEa = ea; }
            else { flush(); runText = ch; runEa = ea; }
          }
          flush();
          return w;
        };
        const pushPiece = (piece: string): void => {
          let runText = '';
          let runEa: boolean | null = null;
          const flush = (): void => {
            if (runText === '') return;
            const pFont = runEa ? (fontEa as string) : font;
            const pFamily = runEa ? (familyEa as string) : family;
            push(runText, pFont, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, { ...segExtras, fontFamily: pFamily });
            runText = '';
          };
          for (const ch of piece) {
            const ea = isEaCh(ch);
            if (runEa === null || ea === runEa) { runText += ch; runEa = ea; }
            else { flush(); runText = ch; runEa = ea; }
          }
          flush();
        };
        // Grapheme-fill runs (Myanmar/Tibetan, #961) have dense per-cluster offsets:
        // use the O(log n) monotone binary-search fit. Dictionary runs keep the
        // negative-spacing-safe full scan.
        const monotone = isGraphemeFillText(token);
        const N = token.length;
        let start = 0;
        while (start < N) {
          const avail = availW();
          let end = fitSeaWordPrefix(token, seaBreaks, start, avail, measureSub, monotone);
          if (end <= start) {
            if (lineW > 0) { newLine(); continue; } // wrap first, retry empty line
            // Empty line, first word wider than the shape: grapheme-safe split.
            const firstWordEnd = seaBreaks.find((b) => b > start) ?? N;
            const firstWord = token.slice(start, firstWordEnd);
            const graphemes = graphemeClusterOffsets(firstWord);
            let g = fitSeaWordPrefix(firstWord, graphemes, 0, avail, measureSub, monotone);
            if (g <= 0) g = graphemes.length > 0 ? graphemes[0] : firstWord.length;
            end = start + g;
          }
          pushPiece(token.slice(start, end));
          start = end;
          if (start < N) newLine();
        }
        continue;
      }

      if (fitsW(tokW)) {
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        if (isWhitespace) hasWhitespaceOnLine = true;
      } else if (isWhitespace) {
        if (lineW > 0) newLine();
      } else if (tokW > lineMaxW()) {
        if (lineW > 0) newLine();
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (!fitsW(chW) && lineW > 0) newLine();
          push(ch, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        }
      } else if (!hasWhitespaceOnLine) {
        // No whitespace yet on this line — wrapping here would tear an
        // unbroken sequence of non-whitespace text (e.g. "YoY+11.9%" split
        // across mixed-size runs). Office never breaks mid-sequence in that
        // case; it lets the shape overflow and relies on spAutoFit / lIns to
        // size the bbox correctly. Match that behavior.
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
      } else {
        // UAX #14 segment-boundary glue: LB13 keeps a non-starter with the word
        // before it; the shared no-break pair predicate (LB14/LB23/LB23a/LB24/
        // LB25/LB28/LB30) keeps proven no-break seams together across a
        // formatting run seam. CJK and SEA tokens have already taken their
        // dedicated paths. Move the trailing word down to the previous real
        // opportunity.
        const previousText = currentLine.segments.at(-1)?.text ?? '';
        const firstCp = token.codePointAt(0);
        const previousChar = [...previousText].at(-1);
        const prevCp = previousChar?.codePointAt(0);
        const immediateBoundary =
          /\S$/u.test(previousText) &&
          /^\S/u.test(token) &&
          prevCp !== 0x200b &&
          firstCp !== 0x200b;
        const lb13Glued =
          firstCp !== undefined &&
          DEFAULT_KINSOKU_RULES.lineStartForbidden.has(firstCp) &&
          immediateBoundary;
        // SEA (Thai/Lao/Khmer) tailoring wins over the LB1 SA→AL default on
        // BOTH sides: a preceding SEA segment exposes a dictionary boundary
        // that the pair predicate must not suppress (mirror the DOCX
        // buildSegments guard, which checks prev AND cur). A SEA `token`
        // already took the dedicated SEA branch above.
        const uax14Glued =
          prevCp !== undefined &&
          firstCp !== undefined &&
          immediateBoundary &&
          !containsSeaScript(previousText) &&
          !containsSeaScript(token) &&
          isUax14NoBreakPair(prevCp, firstCp);
        const glued = lb13Glued || uax14Glued;
        if (!(glued && retractTrailingWord())) newLine();
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
      }
    }
  }

  // Always emit the last (possibly empty) line
  lines.push(currentLine);

  return lines;
}

// ===== Element renderers =====

async function renderBackground(
  ctx: CanvasRenderingContext2D,
  fill: Fill | null,
  canvasW: number,
  canvasH: number,
  scale: number,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
) {
  // ECMA-376 §20.1.8.14 — image (blipFill) background. Paint an opaque white
  // base first so a partially transparent image (alphaModFix) composites over
  // white, and so a decode failure still leaves a defined background.
  if (fill && fill.fillType === 'image') {
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, canvasW, canvasH);
    // The lazy pipeline always emits imagePath + mimeType for blip fills; bail
    // to the white base if either the path or the byte source is missing.
    if (!fill.imagePath || !fill.mimeType || !fetchImage) return;
    try {
      // Size the metafile raster from the fill box (canvasW/H are CSS px;
      // scale is px-per-EMU, so px/scale = EMU, /PT_TO_EMU = pt).
      // §20.1.8.23 duotone recolour on the raster blip (issue #889): route
      // through the shared duotone cache (keyed by path + colours). No duotone ⇒
      // this is exactly the former `getCachedBitmapByPath` decode, byte-identical.
      const bitmap = await getCachedDuotoneBitmapByPath(
        fill.imagePath,
        fill.mimeType,
        fill.duotone,
        fetchImage,
        {
          widthPt: canvasW / scale / PT_TO_EMU,
          heightPt: canvasH / scale / PT_TO_EMU,
        },
      );
      // A null bitmap (unsupported metafile, e.g. true EMF) → keep the white
      // base painted above as the fallback, exactly like a decode failure.
      if (!bitmap) return;
      ctx.save();
      // Clip to the slide rectangle so overscan (negative insets) or tile
      // bleed is cropped at the slide edge rather than spilling onto
      // neighbouring content.
      ctx.beginPath();
      ctx.rect(0, 0, canvasW, canvasH);
      ctx.clip();
      if (fill.alpha != null) ctx.globalAlpha = fill.alpha;
      if (fill.tile) {
        // §20.1.8.58 — tiled placement: repeat the blip at its native size.
        paintTiledBackground(ctx, bitmap, fill.tile, canvasW, canvasH, scale);
      } else {
        // §20.1.8.56 stretch into the destination rect from the §20.1.8.30
        // fillRect insets. l/t are left/top insets, r/b are right/bottom
        // insets, so the destination spans [l, 1-r] × [t, 1-b] of the box;
        // negative edges overscan past the box.
        const fr = fill.fillRect ?? {};
        const l = fr.l ?? 0;
        const t = fr.t ?? 0;
        const r = fr.r ?? 0;
        const b = fr.b ?? 0;
        const dx = l * canvasW;
        const dy = t * canvasH;
        const dw = canvasW * (1 - l - r);
        const dh = canvasH * (1 - t - b);
        ctx.drawImage(bitmap, dx, dy, dw, dh);
      }
      ctx.restore();
    } catch {
      // Decode failed — the white base painted above remains as the fallback.
    }
    return;
  }
  const bg = resolveShapeFill(fill, ctx, 0, 0, canvasW, canvasH);
  ctx.fillStyle = bg ?? '#FFFFFF';
  ctx.fillRect(0, 0, canvasW, canvasH);
}

/**
 * EMU per pixel at 96 DPI. A blip's intrinsic pixel size is interpreted at
 * 96 DPI to get its native EMU size, matching how PowerPoint sizes a tile.
 * 914400 EMU / inch ÷ 96 px / inch = 9525 EMU / px.
 */
const EMU_PER_PX_96 = 9525;

/**
 * Compute the anchor (origin) of the first tile inside the fill box for a
 * given §20.1.8.41 ST_RectAlignment value. The returned point is where the
 * tile grid is registered; tx/ty then shift it further. The pattern is then
 * phase-shifted so a tile edge passes through this anchor.
 *
 * - `tl` → box origin (0,0); the first tile's top-left sits at the box origin.
 * - `ctr` → box centre; a tile is centred in the box.
 * - `br` → box bottom-right corner; a tile's bottom-right sits there.
 * etc. The anchor is expressed as the position (px) of the tile's top-left
 * for the corresponding alignment, before tx/ty.
 */
export function tileAnchorOffset(
  algn: string,
  boxW: number,
  boxH: number,
  tileW: number,
  tileH: number,
): { ax: number; ay: number } {
  // Horizontal: tl/l/bl = left, t/ctr/b = centre, tr/r/br = right.
  let ax: number;
  if (algn === 't' || algn === 'ctr' || algn === 'b') {
    ax = (boxW - tileW) / 2;
  } else if (algn === 'tr' || algn === 'r' || algn === 'br') {
    ax = boxW - tileW;
  } else {
    ax = 0; // tl, l, bl (and any unknown) anchor left.
  }
  // Vertical: tl/t/tr = top, l/ctr/r = middle, bl/b/br = bottom.
  let ay: number;
  if (algn === 'l' || algn === 'ctr' || algn === 'r') {
    ay = (boxH - tileH) / 2;
  } else if (algn === 'bl' || algn === 'b' || algn === 'br') {
    ay = boxH - tileH;
  } else {
    ay = 0; // tl, t, tr (and any unknown) anchor top.
  }
  return { ax, ay };
}

/**
 * Paint a tiled blip background (ECMA-376 §20.1.8.58 CT_TileInfoProperties).
 *
 * The blip repeats at its native pixel size (interpreted at 96 DPI → EMU)
 * scaled by sx/sy and the slide `scale`. `flip` mirrors alternate tiles, which
 * we pre-compose into a 2×2 "super-tile" so a plain `repeat` pattern reproduces
 * the mirror cadence. `algn` registers the grid against a box corner/edge and
 * tx/ty add a further EMU offset. The whole pattern is phase-shifted via a
 * `DOMMatrix` translate on the pattern transform.
 *
 * The caller has already clipped to the slide box and set globalAlpha.
 */
function paintTiledBackground(
  ctx: CanvasRenderingContext2D,
  bitmap: ImageBitmap,
  tile: TileInfo,
  canvasW: number,
  canvasH: number,
  scale: number,
): void {
  // Native tile size in slide px: image px → EMU @96dpi → × sx/sy → × scale.
  const tileW = bitmap.width * EMU_PER_PX_96 * tile.sx * scale;
  const tileH = bitmap.height * EMU_PER_PX_96 * tile.sy * scale;
  if (!(tileW > 0) || !(tileH > 0)) return;

  const flipX = tile.flip === 'x' || tile.flip === 'xy';
  const flipY = tile.flip === 'y' || tile.flip === 'xy';

  // Build the repeating cell. Without flip it is one tile; with flip it is a
  // 2×2 block whose neighbours are mirrored, so the seam between repeats is a
  // mirror line (PowerPoint's tile-flip behaviour).
  const cellW = tileW * (flipX ? 2 : 1);
  const cellH = tileH * (flipY ? 2 : 1);
  const aux = createAuxCanvas(cellW, cellH);
  if (!aux) return;
  const actx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!actx) return;

  const drawCell = (cx: number, cy: number, mx: boolean, my: boolean) => {
    actx.save();
    actx.translate(cx + (mx ? tileW : 0), cy + (my ? tileH : 0));
    actx.scale(mx ? -1 : 1, my ? -1 : 1);
    actx.drawImage(bitmap, 0, 0, tileW, tileH);
    actx.restore();
  };
  // Top-left tile is always un-mirrored; mirror the X/Y neighbours.
  drawCell(0, 0, false, false);
  if (flipX) drawCell(tileW, 0, true, false);
  if (flipY) drawCell(0, tileH, false, true);
  if (flipX && flipY) drawCell(tileW, tileH, true, true);

  const pattern = ctx.createPattern(aux as unknown as CanvasImageSource, 'repeat');
  if (!pattern) return;

  // Phase: register the grid against the alignment anchor, then add tx/ty.
  const { ax, ay } = tileAnchorOffset(tile.algn, canvasW, canvasH, tileW, tileH);
  const px = ax + emuToPx(tile.tx, scale);
  const py = ay + emuToPx(tile.ty, scale);
  // `setTransform` exists where DOMMatrix is available (browser + skia-canvas).
  // The translate aligns a cell origin with (px, py); `repeat` covers the box.
  if (typeof pattern.setTransform === 'function' && typeof DOMMatrix !== 'undefined') {
    pattern.setTransform(new DOMMatrix().translateSelf(px, py));
    ctx.fillStyle = pattern;
    ctx.fillRect(0, 0, canvasW, canvasH);
  } else {
    // Fallback: bake the phase into the fill origin by translating the context.
    ctx.save();
    ctx.translate(px, py);
    ctx.fillStyle = pattern;
    ctx.fillRect(-px, -py, canvasW, canvasH);
    ctx.restore();
  }
}

function applyShadow(ctx: CanvasRenderingContext2D, shadow: Shadow | null, scale: number) {
  if (!shadow) return;
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dist = emuToPx(shadow.dist, scale);
  ctx.shadowColor = hexToRgba(shadow.color, shadow.alpha);
  ctx.shadowBlur = emuToPx(shadow.blur, scale);
  ctx.shadowOffsetX = Math.cos(dirRad) * dist;
  ctx.shadowOffsetY = Math.sin(dirRad) * dist;
}

/**
 * ECMA-376 §20.1.8.17 — apply glow as a coloured halo. The Canvas shadow
 * primitive is a fine fit: zero offset + the glow's blur radius produces a
 * symmetric coloured blur centred on every drawn pixel.
 */
function applyGlow(ctx: CanvasRenderingContext2D, glow: Glow | null | undefined, scale: number) {
  if (!glow) return;
  ctx.shadowColor = hexToRgba(glow.color, glow.alpha);
  ctx.shadowBlur = emuToPx(glow.radius, scale);
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

function clearShadow(ctx: CanvasRenderingContext2D) {
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

// ── WordArt paired-edge warp: piecewise-affine strip subdivision ─────────────
//
// Each glyph is drawn as a stack of vertical STRIPS, every strip re-issuing
// `fillText` under its own affine with a clip rect bounding the strip's share
// of the glyph advance. Drawing VECTOR text per strip (instead of slicing a
// pre-rasterised glyph bitmap, as an earlier revision did) rasterises the
// glyph at the strip's final device transform, which kills three
// bitmap-transfer artifact classes a real browser exposed: blur from
// re-magnifying a raster (vScale stretches up to ~4×), glyph parts cut by the
// offscreen slab rectangle, and detached hairlines from stretched source-edge
// antialiasing. It also needs no auxiliary canvas at all, so the draw works
// identically on the main thread, in a worker, and under headless node.

// BASE strip width in DEVICE px — the coarse first guess only; the adaptive
// deviation loop below is what actually guarantees accuracy. Width alone bounds
// just the chord-vs-curve error of the BASELINE curve (∝ width², sub-pixel at
// 8 px for every preset), but NOT the error away from the baseline: a slab
// corner at distance d from the baseline anchor is additionally misplaced by
// the angle/vScale mismatch across the strip (≈ d·Δangle + |y|·ΔvScale), which
// at 8 px width is measured at ~0.7 px (Inflate) to ~0.2 px (Wave1/Chevron) at
// the baseline but grows to several px at the em-box extremes (Wave1 ≈2.4,
// Button ≈5.9) and tens of px on high-curvature ring/pour families
// (CirclePour ≈11, CanUp ≈44 at em-box edges) — the strips fan apart into
// radial slivers. Hence the width criterion is only the starting k; the
// deviation-driven refinement below subdivides until every painted corner is
// within WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX of the exact envelope map.
const WARP_STRIP_DEVICE_PX = 8;
// Extend each strip's CLIP rect by 1 device px into its neighbour on each
// side. Two abutting clip edges antialiased in separate fill passes leave the
// classic shared-edge coverage dip (each pass covers the boundary pixel
// partially; src-over of two ~50 % coverages of colour C over white composes
// to ~0.75·C, a faint light seam through every glyph) — and the two edges can
// additionally diverge by up to 2·budget at the ink extremes. Overlapping the
// clips makes strip i paint the boundary band at FULL coverage before strip
// i+1's partial-coverage edge lands on top; painting the same opaque colour
// over itself is idempotent, so the seam vanishes. That idempotence only holds
// for an OPAQUE composite: a translucent fill (alpha < 1, from an 8-digit
// RRGGBBAA run colour or an inherited shape opacity) composes the 2-px band
// TWICE and darkens it (issue #879 — a visible pinstripe/moiré). The translucent
// branch of drawWarpedGlyphStrips fixes that by painting the strips OPAQUE into a
// per-glyph device-resolution layer (where the overlap is idempotent again) and
// compositing the layer ONCE with the effective alpha. The overlap is symmetric,
// so it does not bias the mapping.
const WARP_STRIP_OVERLAP_DEVICE_PX = 1;
// Max deviation (device px) of any painted slab corner from the EXACT envelope
// map, enforced by the adaptive loop. Derivation: the true map places the
// glyph's vertical ruling at flat-x b onto Q(b, y) = A(u_b)·(0, y), where A(u)
// is the envelope's local frame at u (anchor T(u)+bandFrac·gap, advance-axis
// angle, gap shear/scale — exactly what warpGlyphTransform returns). A strip
// centred at c paints that same ruling at S(b, y) = A(u_c)·(b−c, y). The
// deviation |S−Q| at the slab's four corners (both strip edges × slab
// top/bottom) is the exact per-strip error of the piecewise-affine map — it
// contains ALL mismatch components at once: rotational fan-out d·Δangle
// (dominant on rings/pours, whose tangent sweeps up to ~345°), vertical-scale
// slide |y|·ΔvScale (dominant on textButton, whose mean-tangent angle is
// constant 0 so a pure angle criterion would miss it), and the baseline chord
// error. Two adjacent strips each deviate ≤ budget from the SAME true ruling at
// their shared boundary, so their mutual divergence is ≤ 2·budget = the
// 2 device px their overlapping clips share — i.e. budget 1 is the largest
// value the clip overlap provably covers. Every quantity is geometric; no
// preset branch.
const WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX = 1;
// Hard cap on strips per glyph, bounding draw cost for extreme geometry (one
// glyph sweeping a large arc of a ring at high em-height). At the cap the
// residual deviation is reported by the loop's last measurement and the seam
// stays proportional (a full-ring glyph at ~250 device px em-height measures a
// few px — degraded but bounded, and localized to that pathological glyph).
const WARP_STRIP_MAX_PER_GLYPH = 256;

/**
 * Draw one glyph across a PAIRED-EDGE warp envelope as a stack of vertical
 * strips (piecewise-affine approximation of the envelope map, ECMA-376
 * §20.1.9.19).
 *
 * The #866 per-glyph transform samples the envelope's local Jacobian at a SINGLE
 * u (the glyph centre) and applies one affine (rotate + shear + non-uniform
 * scale) to the whole glyph. That is exact only where the map is locally affine;
 * on Inflate/Deflate the vertical stretch `vScale` varies WITHIN a glyph (the
 * top/bottom edges are mirror-symmetric so the gap stays vertical — shear 0,
 * angle 0 — but the gap MAGNITUDE peaks at the centre and shrinks toward the
 * ends), which a single affine cannot represent: it leaves the glyph uniformly
 * scaled instead of stretched more where the envelope is taller. PowerPoint maps
 * the glyph OUTLINE continuously through `P(u,v) = (1−v)·T(u) + v·B(u)`, so a
 * wide glyph leans/stretches asymmetrically across its own width.
 *
 * Subdividing the glyph into narrow strips and mapping each strip by the Jacobian
 * at ITS centre-u converges to that continuous map as the strip count grows (the
 * general solution — no preset special-casing). The count is chosen ADAPTIVELY:
 * from a base width criterion, strips double until every painted corner deviates
 * from the exact envelope map by at most 1 device px (see
 * WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX), so high-curvature envelopes (rings,
 * pours — tangent sweep up to ~345°) subdivide as finely as they need while flat
 * ones stay coarse.
 *
 * Each strip re-issues `fillText` under its own affine, clipped to the strip's
 * share of the glyph advance. The clip rect is expressed in GLYPH-LOCAL
 * (flat-advance) coordinates, so adjacent strips cut the glyph at the same
 * flat-x boundary; the text itself is rasterised by the canvas at the strip's
 * final device-space transform — full device resolution, no intermediate
 * bitmap. The clip is effectively unbounded vertically (the subdivision is
 * horizontal only) and the OUTER edges of a glyph's first/last strips are
 * unbounded outward, so side-bearing overhang ink (italic tails, swashes)
 * rides the end strips instead of being cut at the advance box.
 *
 * `chW` is the glyph's flat advance (css px, incl. letter-spacing `ls`);
 * `flatX0` is the glyph's flat left edge along the whole line (css px). The
 * remaining args mirror the per-glyph draw: `hScale` stretches flat ink to the
 * box width, `warpBoxH`/`bandFrac` drive `warpGlyphTransform`, `boxX`/`boxY`
 * are the shape origin, `totalW`/`followScale` map flat-x to the envelope
 * fraction u. `ctx.font` / `ctx.fillStyle` must already be set for this glyph
 * (save/restore around each strip preserves them).
 */
function drawWarpedGlyphStrips(
  ctx: CanvasRenderingContext2D,
  ch: string,
  ls: number,
  chW: number,
  devScale: number,
  env: WarpEnvelope,
  flatX0: number,
  totalW: number,
  followScale: number,
  hScale: number,
  warpBoxH: number,
  bandFrac: number,
  boxX: number,
  boxY: number,
  color: string,
): void {
  if (chW <= 0) return;
  // Ink extremes about the baseline (css px), from real metrics — this is where
  // the deviation budget must hold, since that is where ink actually lands.
  // Fallbacks keep degenerate metrics safe (mirrors the flat renderer).
  const m = ctx.measureText(ch);
  const asc = m.actualBoundingBoxAscent > 0 ? m.actualBoundingBoxAscent : chW;
  const desc = m.actualBoundingBoxDescent > 0 ? m.actualBoundingBoxDescent : chW * 0.25;
  // Horizontal ink extremes about the pen origin (css px). Used only to bound the
  // translucent compositing layer (issue #879) tightly around real ink, so
  // side-bearing overhang on the outward-unbounded end strips still fits.
  const inkL = m.actualBoundingBoxLeft > 0 ? m.actualBoundingBoxLeft : 0;
  const inkR = m.actualBoundingBoxRight > 0 ? m.actualBoundingBoxRight : chW;

  // ── Adaptive strip count ──────────────────────────────────────────────────
  // Start from the width criterion (one strip per WARP_STRIP_DEVICE_PX of the
  // glyph's WARPED advance, chW·hScale·devScale), then DOUBLE the count until
  // every painted ink corner is within the deviation budget of the exact
  // envelope map (see WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX for the formula and
  // why the budget ties to the clip overlap). Deviation from the smooth part of
  // the envelope halves with each doubling (it is ∝ the strip's u-span to first
  // order), so the loop converges in a few steps for arcs of any curvature —
  // this is what keeps high-curvature ring/pour envelopes (tangent sweep up to
  // ~345°) contiguous instead of fanning into radial slivers. If a doubling
  // fails to reduce the deviation by ≥ 25 %, the residual sits at a C0 corner of
  // the envelope itself (e.g. the chevron apex): the true map genuinely creases
  // there, refinement can only narrow the affected band — keep the finer
  // partition and stop rather than spinning to the cap.
  const warpedAdvDev = chW * hScale * devScale;
  let k = Math.min(
    WARP_STRIP_MAX_PER_GLYPH,
    Math.max(1, Math.round(warpedAdvDev / WARP_STRIP_DEVICE_PX)),
  );
  const build = (count: number): WarpStrip[] =>
    buildWarpStrips(count, env, chW, flatX0, totalW, followScale, warpBoxH, bandFrac);
  let strips = build(k);
  let dev = maxWarpStripDeviationDev(
    strips, env, flatX0, totalW, followScale, warpBoxH, bandFrac, hScale, devScale, -asc, desc,
  );
  while (dev > WARP_STRIP_DEVIATION_BUDGET_DEVICE_PX && k < WARP_STRIP_MAX_PER_GLYPH) {
    const k2 = Math.min(WARP_STRIP_MAX_PER_GLYPH, k * 2);
    const strips2 = build(k2);
    const dev2 = maxWarpStripDeviationDev(
      strips2, env, flatX0, totalW, followScale, warpBoxH, bandFrac, hScale, devScale, -asc, desc,
    );
    if (dev2 >= dev * 0.75) {
      // Not converging: C0 crease in the envelope (see above). The finer
      // partition still localizes the crease to a narrower strip — keep it.
      strips = strips2;
      break;
    }
    k = k2;
    strips = strips2;
    dev = dev2;
  }

  // Vertical clip extent: the subdivision is horizontal only, so the clip's
  // sole job is to bound x. ±10⁴ css px is four orders of magnitude beyond any
  // glyph's extent (so no ink is ever cut in y), kept finite only so the
  // rasterizer's float precision on the boundary line stays sub-thousandth-px.
  const CLIP_Y = 1e4;
  // Clip overlap in glyph-local flat units: the local x axis is scaled by
  // hScale (then the canvas dpr) on the way to device px, so this makes the
  // painted overlap WARP_STRIP_OVERLAP_DEVICE_PX device px wide.
  const ov = WARP_STRIP_OVERLAP_DEVICE_PX / (hScale * devScale);

  const last = strips.length - 1;
  // Glyph-local clip interval [x0, x1] of strip `i` (flat css px, origin at the
  // strip centre): interior boundaries sit at the shared s0/s1 (± the overlap);
  // the outer edges of the first/last strips are unbounded so side-bearing
  // overhang ink is not cut.
  const stripClipX0 = (i: number, s0: number, centre: number): number =>
    i === 0 ? -CLIP_Y : s0 - centre - ov;
  const stripClipX1 = (i: number, s1: number, centre: number): number =>
    i === last ? CLIP_Y : s1 - centre + ov;

  // Paint the whole strip stack onto `target` in `fillStyle`. The transform
  // chain per strip is the single-affine per-glyph draw anchored at the STRIP
  // centre so vScale/shear/angle track this slice's u.
  const paintStrips = (target: CanvasRenderingContext2D, fillStyle: string): void => {
    target.fillStyle = fillStyle;
    for (let i = 0; i <= last; i++) {
      const { s0, s1, g } = strips[i];
      const centre = (s0 + s1) / 2;
      target.save();
      target.translate(boxX + g.x, boxY + g.y);
      target.rotate(g.angle);
      if (g.shear !== 0) target.transform(1, 0, g.shear, 1, 0, 0);
      if (hScale !== 1 || g.vScale !== 1) target.scale(hScale, g.vScale);
      target.beginPath();
      const x0 = stripClipX0(i, s0, centre);
      const x1 = stripClipX1(i, s1, centre);
      target.rect(x0, -CLIP_Y, x1 - x0, 2 * CLIP_Y);
      target.clip();
      // Pen origin sits at local −centre; the ls/2 shift centres the ink inside
      // its letter-spaced advance, matching the flat draw's origin convention.
      target.fillText(ch, -centre + ls / 2, 0);
      target.restore();
    }
  };

  // ── Composite path selection (issue #879) ──────────────────────────────────
  // The overlapping strip clips are idempotent only for a FULLY OPAQUE composite.
  // Effective opacity = the fill's own alpha × the inherited ctx.globalAlpha
  // (shape opacity). When both are 1 the direct draw is exact and byte-identical
  // to the pre-#879 path — the overwhelmingly common WordArt case, zero
  // allocation. Otherwise the 1-device-px overlap band would double-compose and
  // darken, so route through the layer below.
  const fillAlpha = rgbaAlpha(color);
  const destAlpha = typeof ctx.globalAlpha === 'number' ? ctx.globalAlpha : 1;
  if (fillAlpha >= 1 && destAlpha >= 1) {
    paintStrips(ctx, color);
    return;
  }
  if (fillAlpha <= 0 || destAlpha <= 0) return; // fully transparent — nothing to draw

  // Translucent: paint the strips OPAQUE into a per-glyph device-resolution layer
  // (so the overlap band is idempotent again) and composite the layer ONCE with
  // the effective alpha. The layer is rasterised under the SAME device transform
  // as the live canvas, so there is no re-magnification of any raster — the fix
  // the issue calls for. Falls back to the direct draw (today's slightly darker
  // band) when no auxiliary canvas or live transform is available (a headless
  // mock ctx), which is never pixel-verified anyway.
  const base = typeof ctx.getTransform === 'function' ? ctx.getTransform() : null;
  if (!base) {
    paintStrips(ctx, color);
    return;
  }

  // Device-space AABB of the actual ink: per strip, the visible ink is the
  // measured ink rectangle intersected with the strip's clip interval, mapped
  // through the strip affine and the live (css→device) transform. Bounding by
  // measured ink — not the ±CLIP_Y clip — keeps the layer glyph-tight.
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (let i = 0; i <= last; i++) {
    const { s0, s1, g } = strips[i];
    const centre = (s0 + s1) / 2;
    const textX = -centre + ls / 2;
    const x0 = Math.max(stripClipX0(i, s0, centre), textX - inkL);
    const x1 = Math.min(stripClipX1(i, s1, centre), textX + inkR);
    if (x1 <= x0) continue; // no ink in this strip's clip
    for (const [lx, ly] of [
      [x0, -asc],
      [x1, -asc],
      [x0, desc],
      [x1, desc],
    ] as const) {
      const p = mapWarpLocalPoint(g, hScale, lx, ly);
      const cx = boxX + p.x;
      const cy = boxY + p.y;
      const dx = base.a * cx + base.c * cy + base.e;
      const dy = base.b * cx + base.d * cy + base.f;
      if (dx < minX) minX = dx;
      if (dx > maxX) maxX = dx;
      if (dy < minY) minY = dy;
      if (dy > maxY) maxY = dy;
    }
  }
  if (!(maxX > minX && maxY > minY)) return; // no ink mapped — nothing to composite

  const pad = 2; // device px, for AA/rounding slack around the ink AABB
  const originX = Math.floor(minX - pad);
  const originY = Math.floor(minY - pad);
  const layerW = Math.ceil(maxX + pad) - originX;
  const layerH = Math.ceil(maxY + pad) - originY;
  const aux = createAuxCanvas(layerW, layerH);
  const auxCtx = aux ? (aux.getContext('2d') as CanvasRenderingContext2D | null) : null;
  if (!aux || !auxCtx) {
    paintStrips(ctx, color);
    return;
  }
  // The layer's own transform is the live transform translated by −origin (a
  // pure device-space shift, so it just subtracts origin from the base
  // translation — no DOMMatrix multiply needed), placing device pixels on the
  // same grid as the main canvas but offset into the layer.
  auxCtx.font = ctx.font;
  auxCtx.textAlign = 'left';
  auxCtx.textBaseline = 'alphabetic';
  auxCtx.setTransform(base.a, base.b, base.c, base.d, base.e - originX, base.f - originY);
  paintStrips(auxCtx, opaqueRgba(color));

  // Composite the opaque layer once at the effective alpha, at identity so the
  // (already device-oriented) layer blits 1:1. save/restore preserves the live
  // transform, clip, and globalAlpha for the next glyph.
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = destAlpha * fillAlpha;
  ctx.drawImage(aux, originX, originY);
  ctx.restore();
}

/** Alpha (0..1) of an `rgb()/rgba()` colour string. The pptx text pipeline emits
 *  colours via {@link hexToRgba} (`rgba(r,g,b,a)`); a 3-component `rgb()` or any
 *  unparseable value is treated as opaque so it never routes into the #879 layer. */
function rgbaAlpha(color: string): number {
  const m = /^rgba?\(\s*[\d.]+\s*,\s*[\d.]+\s*,\s*[\d.]+\s*,\s*([\d.]+)\s*\)$/i.exec(color);
  if (!m) return 1;
  const a = parseFloat(m[1]);
  return Number.isFinite(a) ? Math.min(1, Math.max(0, a)) : 1;
}

/** Drop the alpha channel of an `rgb()/rgba()` colour, yielding an opaque
 *  `rgb(r,g,b)`. Used to paint the #879 compositing layer at full opacity. */
function opaqueRgba(color: string): string {
  const m = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/i.exec(color);
  return m ? `rgb(${m[1]}, ${m[2]}, ${m[3]})` : color;
}

/** One strip of a glyph: flat sub-range `[s0, s1]` (css px from the glyph's flat
 *  left edge / pen origin) and the envelope frame sampled at its centre-u. */
interface WarpStrip {
  s0: number;
  s1: number;
  g: WarpGlyphTransform;
}

/** Uniform-in-u partition of a glyph's advance into `k` strips, each carrying
 *  the envelope frame at its own centre-u. */
function buildWarpStrips(
  k: number,
  env: WarpEnvelope,
  chW: number,
  flatX0: number,
  totalW: number,
  followScale: number,
  warpBoxH: number,
  bandFrac: number,
): WarpStrip[] {
  const strips: WarpStrip[] = new Array(k);
  for (let j = 0; j < k; j++) {
    const s0 = (j / k) * chW;
    const s1 = ((j + 1) / k) * chW;
    const u = ((flatX0 + (s0 + s1) / 2) / totalW) * followScale;
    strips[j] = { s0, s1, g: warpGlyphTransform(env, u, warpBoxH, bandFrac) };
  }
  return strips;
}

/** Apply a strip's affine (scale → shear → rotate → translate, the same chain
 *  the draw path pushes onto the canvas) to a strip-local css point. */
function mapWarpLocalPoint(
  g: WarpGlyphTransform,
  hScale: number,
  x: number,
  y: number,
): { x: number; y: number } {
  const sx = x * hScale;
  const sy = y * g.vScale;
  const hx = sx + g.shear * sy;
  const c = Math.cos(g.angle);
  const s = Math.sin(g.angle);
  return { x: g.x + c * hx - s * sy, y: g.y + s * hx + c * sy };
}

/**
 * Worst-case deviation (device px) of the strips' painted slab corners from the
 * EXACT envelope map — the quantity the adaptive loop drives under the budget.
 *
 * For each strip edge b (a boundary of the flat sub-range), the exact map
 * places the glyph ruling at `Q(b, y) = A(u_b)·(0, y)` — the envelope frame
 * sampled AT the boundary — while the strip paints it at
 * `S(b, y) = A(u_centre)·(b − centre, y)`. `|S − Q|` is evaluated at the slab's
 * top and bottom (`yTop`/`yBot`, css, baseline-relative) for both edges of
 * every strip; the max over all of them bounds every painted pixel's error
 * (corners are the extreme points of an affine image of a rectangle).
 */
function maxWarpStripDeviationDev(
  strips: WarpStrip[],
  env: WarpEnvelope,
  flatX0: number,
  totalW: number,
  followScale: number,
  warpBoxH: number,
  bandFrac: number,
  hScale: number,
  devScale: number,
  yTop: number,
  yBot: number,
): number {
  let max = 0;
  for (const strip of strips) {
    const centre = (strip.s0 + strip.s1) / 2;
    for (const b of [strip.s0, strip.s1]) {
      const uB = ((flatX0 + b) / totalW) * followScale;
      const gB = warpGlyphTransform(env, uB, warpBoxH, bandFrac);
      for (const y of [yTop, yBot]) {
        const q = mapWarpLocalPoint(gB, hScale, 0, y);
        const p = mapWarpLocalPoint(strip.g, hScale, b - centre, y);
        const d = Math.hypot(p.x - q.x, p.y - q.y) * devScale;
        if (d > max) max = d;
      }
    }
  }
  return max;
}

/**
 * Render a WordArt text body through its `<a:prstTxWarp>` envelope (ECMA-376
 * §20.1.9.19). Canvas 2D cannot deform glyph outlines, so this is the standard
 * PER-GLYPH approximation: the text is laid out flat (reusing
 * {@link layoutParagraph} so run font/size/colour resolution is identical to the
 * normal path), then each glyph is placed against the preset's flattened
 * envelope via {@link warpGlyphTransform} and drawn with a local
 * translate/rotate/scale.
 *
 * The envelope spans the shape's box (minus insets). Text is fitted to that
 * width: a glyph whose centre sits at horizontal fraction `u` of the total text
 * width is mapped to the envelope point at `u`. Multiple lines stack vertically
 * inside the box, each occupying its own horizontal band of the envelope's
 * height — enough for the common single-line WordArt title.
 *
 * onTextRun is intentionally NOT invoked here: the transparent text-selection
 * overlay assumes axis-aligned run boxes, which per-glyph warped text violates.
 * Warped WordArt is decorative; skipping the overlay leaves it unselectable
 * rather than mis-boxed. This is called only when a warp preset is present, so
 * all non-WordArt text keeps its overlay.
 */
function renderWarpedText(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  preset: string,
  adj: number[],
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  defaultColor: string,
  rc: RenderContext,
): void {
  // The warp envelope spans the FULL shape bounding box, NOT the text rect
  // inset by lIns/tIns/rIns/bIns. The preset guide formulas (ECMA-376
  // §20.1.9.19 / presetTextWarpDefinitions.xml) are written against the
  // shape's `w`/`h` built-ins, and PowerPoint's own PDF output of the warp
  // fixture confirms it: paired-edge warp ink spans 6.04–6.18in of a 6.2in
  // shape (insets would cap it at 6.0) and up to 1.48in of a 1.5in-tall one.
  const boxX = bx;
  const boxY = by;
  const boxW = Math.max(1, bw);
  const boxH = Math.max(1, bh);

  const env = buildWarpEnvelope(preset, adj, boxW, boxH);
  if (!env) return; // unknown preset — nothing to draw (caller already gated)

  const bodyDefaultBold = body.defaultBold ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale;

  // Lay each paragraph out flat at the natural size (no wrap — WordArt fits the
  // shape width itself). Collect every line's segments in order.
  const lines: LayoutLine[] = [];
  for (const para of body.paragraphs) {
    const paraDefaultFontSizePx =
      para.defFontSize != null ? para.defFontSize * PT_TO_EMU * scale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor ? hexToRgba(para.defColor) : defaultColor;
    const laid = layoutParagraph(
      ctx,
      para,
      Infinity, // no wrap: the envelope maps the whole run onto the width
      paraDefaultFontSizePx,
      paraDefaultColor,
      scale,
      0,
      bodyDefaultBold,
      bodyDefaultItalic,
      1.0,
      undefined,
      rc,
      0,
    );
    for (const l of laid) lines.push(l);
  }
  if (lines.length === 0) return;

  ctx.save();
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';

  // Device scale folded into the live transform (dpr · slide scale), so glyph
  // bitmaps for the paired-edge strip path are rasterised at the same pixel
  // density as the canvas — no HiDPI blur. sqrt(|det|) of the transform's linear
  // part is dpr regardless of rotation/flip (same idiom as the scene3d/effect
  // offscreens). Computed LAZILY on the first paired-edge glyph only: single-edge
  // (Follow Path) warps never allocate an offscreen, and some unit tests drive a
  // mock ctx without getTransform, so the single-edge path must never touch it.
  let devScaleCache = -1;
  const getDevScale = (): number => {
    if (devScaleCache >= 0) return devScaleCache;
    const tf = typeof ctx.getTransform === 'function' ? ctx.getTransform() : null;
    const det = tf ? Math.abs(tf.a * tf.d - tf.b * tf.c) : 1;
    devScaleCache = det > 0 ? Math.sqrt(det) : 1;
    return devScaleCache;
  };

  const lineCount = lines.length;
  for (let li = 0; li < lineCount; li++) {
    const line = lines[li];
    // Per-line vertical band [v0, v1] of the envelope's height. One line fills
    // the whole band; multiple lines split it evenly top→bottom.
    const v0 = li / lineCount;
    const v1 = (li + 1) / lineCount;

    // Measure the line: total advance width, max font size, and the line's INK
    // extents (max actualBoundingBox ascent/descent over the segments). The ink
    // box — not the em box — is what PowerPoint normalises onto the envelope.
    let totalW = 0;
    let maxSize = 0;
    let maxA = 0; // ink ascent above the baseline
    let maxD = 0; // ink descent below the baseline
    for (const seg of line.segments) {
      if (seg.math) {
        totalW += seg.math.width;
        maxSize = Math.max(maxSize, seg.sizePx);
        maxA = Math.max(maxA, seg.math.ascent);
        maxD = Math.max(maxD, seg.math.descent);
        continue;
      }
      ctx.font = seg.font;
      const ls = seg.letterSpacingPx ?? 0;
      const m = ctx.measureText(seg.text);
      totalW += m.width + ls * codePointCount(seg.text);
      maxSize = Math.max(maxSize, seg.sizePx);
      if (m.actualBoundingBoxAscent > 0) maxA = Math.max(maxA, m.actualBoundingBoxAscent);
      if (m.actualBoundingBoxDescent > 0) maxD = Math.max(maxD, m.actualBoundingBoxDescent);
    }
    if (totalW <= 0) continue;

    // PowerPoint's WordArt semantics for the envelope (paired-edge) presets:
    // the FLAT text's INK RECTANGLE (totalW × ink height) is first STRETCHED to
    // fill the shape box, and the warp then bends that stretched rectangle
    // between the two edges. Measured against PowerPoint's own PDF of the warp
    // fixture (calibrated print scale), paired-edge ink spans 6.04–6.18in of
    // the 6.2in shape and up to 1.48in of the 1.5in one — i.e. the ink touches
    // the envelope edges. Passing the body height here instead would collapse
    // vScale = gap/boxH ≈ 1 and leave glyphs at their tiny natural size.
    //
    // - Vertical: the line's flat ink box (maxA+maxD px, measured — no 0.8em
    //   heuristic) maps to this line's SHARE of the local gap.
    //   warpGlyphTransform returns vScale = gap/boxHeight, so passing
    //   `inkH / (v1 - v0)` yields gap·(v1−v0)/inkH, and a baseline fraction of
    //   maxA/inkH pins the ink top to T(u) and the ink bottom to B(u).
    // - Horizontal: the glyph itself is widened by boxW/totalW (not just the
    //   spacing). Glyph centres sit Δu·boxW ≈ chW·hScale apart, so scaled
    //   glyphs tile the envelope without gaps or overlaps.
    //
    // Single-edge (arch/circle) presets keep natural glyph size: PowerPoint's
    // "Follow Path" semantics place the text at its NATURAL width along the
    // arc without stretching the glyphs (vScale stays 1 inside
    // warpGlyphTransform); they keep the flat renderer's 0.8 ascent fallback
    // for the baseline drop below the arc. The text also follows the path for
    // only its natural arc-length span from the start (stAng): a glyph's `u`
    // fraction is scaled by naturalWidth/arcLength via `followPathUScale`, so a
    // word narrower than the arc occupies a LEADING segment of the path rather
    // than being scattered over the whole ellipse. Paired-edge presets return
    // scale 1 (they stretch the flat ink box to fill the envelope width).
    const inkH = maxA + maxD > 0 ? maxA + maxD : maxSize;
    const baselineFrac = env.singleEdge ? 0.8 : inkH > 0 ? maxA / inkH : 0.8;
    const hScale = env.singleEdge ? 1 : boxW / totalW;
    const warpBoxH = env.singleEdge ? boxH : inkH / (v1 - v0);
    // Follow Path: fraction of the arc the natural-width text actually spans.
    // 1 for paired-edge (no clamp); ≤1 for arch/circle.
    const followScale = followPathUScale(env, totalW);

    // Walk glyphs left→right. `penW` accumulates the flat advance so each glyph's
    // CENTRE maps to its u fraction; the glyph is drawn at a per-glyph transform.
    let penW = 0;
    for (const seg of line.segments) {
      if (seg.math) {
        // Equations inside WordArt are exotic; advance without warping them.
        penW += seg.math.width;
        continue;
      }
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      const ls = seg.letterSpacingPx ?? 0;
      const chars = [...seg.text];
      for (const ch of chars) {
        const chW = ctx.measureText(ch).width + ls;
        // Blend the per-line vertical band into the baseline fraction so line 2
        // sits below line 1 within the envelope.
        const bandFrac = v0 + baselineFrac * (v1 - v0);

        // Paired-edge (envelope) presets: the envelope map is NON-affine within a
        // glyph (on Inflate/Deflate the vertical stretch varies across the glyph's
        // own width; on waves the slope does), so a single per-glyph affine loses
        // that intra-glyph deformation. Draw the glyph via piecewise-affine STRIPS
        // (each strip a clipped vector fillText under its own affine, sampled at
        // its own centre-u) so the mapping converges to PowerPoint's continuous
        // outline warp (§20.1.9.19). Needs no auxiliary canvas, so it runs
        // unconditionally (main thread, worker, headless node alike). The
        // single-edge (Follow Path) branch below is UNCHANGED (byte-identical):
        // its glyphs are rigidly rotated onto a baseline, already exact per glyph.
        if (!env.singleEdge && chW > 0) {
          drawWarpedGlyphStrips(
            ctx,
            ch,
            ls,
            chW,
            getDevScale(),
            env,
            penW,
            totalW,
            followScale,
            hScale,
            warpBoxH,
            bandFrac,
            boxX,
            boxY,
            seg.color,
          );
          penW += chW;
          continue;
        }

        // Horizontal fraction of THIS glyph's centre along the whole line,
        // scaled by the Follow Path factor so single-edge (arch/circle) text
        // spans only its natural arc length from the start rather than the whole
        // path. `followScale` is 1 for paired-edge presets (unchanged).
        const u = ((penW + chW / 2) / totalW) * followScale;
        const g = warpGlyphTransform(env, u, warpBoxH, bandFrac);
        ctx.save();
        ctx.translate(boxX + g.x, boxY + g.y);
        ctx.rotate(g.angle);
        // Local envelope shear (§20.1.9.19): on a sloped paired edge the em-height
        // axis (gap vector B−T) is not perpendicular to the advance axis, so the
        // glyph must skew — vertical strokes track the gap while the baseline
        // follows the slope. `g.shear` is the horizontal skew in the rotated
        // frame; 0 for flat points and single-edge (arch/circle) presets, so this
        // is a no-op there and leaves Follow Path glyphs rigidly rotated.
        if (g.shear !== 0) ctx.transform(1, 0, g.shear, 1, 0, 0);
        if (hScale !== 1 || g.vScale !== 1) ctx.scale(hScale, g.vScale);
        // Draw the glyph centred on the mapped point: shift left by half its
        // advance so its own centre lands on `u`.
        ctx.fillText(ch, -chW / 2 + ls / 2, 0);
        ctx.restore();
        penW += chW;
      }
    }
  }
  ctx.restore();
}

/**
 * Preset-geometry text rectangle (ECMA-376 §20.1.9.21 `<a:rect>` /
 * presetTextRectangle), as an absolute sub-rect of the shape bbox. PowerPoint
 * lays text out inside this rect (then applies the `lIns/tIns/rIns/bIns`
 * insets), NOT in the full bounding box. For arrows the text rect is the shaft
 * (left of the arrowhead), so centered text sits in the body — without this a
 * `rightArrow`'s label drifts right into the arrowhead. Returns null when the
 * geometry's text rect is the whole bbox (the common case).
 *
 * Arrowhead length uses `ss = min(w, h)` per the spec's `dx1 = ss * adj2`.
 */
function presetTextRect(
  geom: string,
  x: number, y: number, w: number, h: number,
  adj1?: number | null, adj2?: number | null,
): { tx: number; ty: number; tw: number; th: number } | null {
  const ss = Math.min(w, h);
  switch (geom) {
    case 'rightarrow':
    case 'leftarrow': {
      const a1 = Math.min(Math.max(adj1 ?? 50000, 0), 100000); // shaft height fraction
      const a2 = Math.min(Math.max(adj2 ?? 50000, 0), 100000); // arrowhead length fraction
      const dx = (ss * a2) / 100000;        // arrowhead length (ss-based, ECMA dx1)
      const dy = (h * a1) / 200000;          // shaft half-height about the center
      const ty = y + h / 2 - dy;
      const th = 2 * dy;
      const tw = Math.max(0, w - dx);
      return geom === 'rightarrow'
        ? { tx: x, ty, tw, th }              // shaft on the left, arrowhead on the right
        : { tx: x + dx, ty, tw, th };        // leftarrow: arrowhead on the left
    }
    case 'roundrect': {
      // Text rect inset from the rounded corners: il = radius * (1 - 1/√2),
      // radius = ss * adj / 100000 (ECMA roundRect adj default 16667).
      const a = Math.min(Math.max(adj1 ?? 16667, 0), 100000) as number;
      const il = (ss * a) / 100000 * (1 - 1 / Math.SQRT2);
      return { tx: x + il, ty: y + il, tw: Math.max(0, w - 2 * il), th: Math.max(0, h - 2 * il) };
    }
    default:
      return null;
  }
}

/**
 * The shape-level default text colour passed to {@link renderTextBody} (used
 * by runs whose colour resolves to null). A style-derived `defaultTextColor`
 * (p:style > fontRef) always wins; the synthetic SmartArt fallback shape —
 * which never carries one — takes the slide's contrast-aware default from
 * {@link RenderContext.smartArtFallbackTextColor} so its data-model text is
 * legible on a dark background (issue #805). All other shapes: null (theme
 * default applies downstream).
 */
function shapeDefaultTextColor(el: ShapeElement, rc: RenderContext): string | null {
  if (el.defaultTextColor) return hexToRgba(el.defaultTextColor);
  if (rc.smartArtFallbackTextColor != null && isSmartArtFallbackShape(el)) {
    return rc.smartArtFallbackTextColor;
  }
  return null;
}

function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number, themeDefaultColor = '#000000', slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, onTextRun?: TextRunCallback, fetchImage?: FetchImage) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);
  const shapeTextRunCallback: TextRunCallback | undefined = onTextRun && el.id !== undefined
    ? (run) => onTextRun({ ...run, shapeId: el.id })
    : onTextRun;

  // anchor="b" + h=0: shape grows upward from y; render stroke as bottom border,
  // then let renderTextBody handle positioning.
  if (h === 0 && el.textBody?.verticalAnchor === 'b') {
    if (el.stroke) {
      ctx.save();
      applyStroke(ctx, el.stroke, scale);
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.stroke();
      ctx.restore();
    }
    if (el.textBody) {
      const defaultTextColor = shapeDefaultTextColor(el, rc);
      renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor, slideNumber, rc, shapeTextRunCallback, false, fetchImage);
    }
    return;
  }

  // ── scene3d camera projection for p:sp (ECMA-376 §20.1.5.5, Phase B) ──────
  // A text-bearing shape with a non-identity 3-D camera is rendered to a local
  // offscreen (fill + stroke + TEXT), optionally bevel-shaded, then warped onto
  // the live ctx through the camera homography — the same pipeline the picture
  // path uses. The element's 2-D rotation/flip stays on the live ctx so the warp
  // composes inside it ("scene3d first, then xfrm", §20.1.5.5).
  //
  // LIMITATION (documented, not silent): the HTML text-selection overlay tracks
  // glyphs by their un-projected layout rect and CSS transforms; it cannot
  // follow the per-pixel perspective warp applied here. So a projected shape's
  // text is drawn into the bitmap but EXCLUDED from the selection overlay
  // (onTextRun is suppressed below). See README "Bevel / 3D" note.
  const spScene3d = el.scene3d && isScene3dNonIdentity(el.scene3d.camera) ? el.scene3d : null;
  if (spScene3d && w > 0 && h > 0) {
    const tf = ctx.getTransform();
    const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
    const ctxDevScale = det > 0 ? Math.sqrt(det) : 1;
    const bevels = buildBevelInputs(
      el.sp3d as Sp3dLike | undefined,
      el.scene3d?.lightRig as LightRigLike | undefined,
      (el.sp3d as { prstMaterial?: string } | undefined)?.prstMaterial,
      scale,
      ctxDevScale,
    );
    const extrusion = buildExtrusion(
      el.sp3d as Sp3dLike | undefined,
      spScene3d.camera,
      w,
      h,
      scale,
      ctxDevScale,
    );
    // Apply the element's own rotation/flip on the live ctx; the warp composes
    // inside it.
    ctx.save();
    if (el.rotation !== 0 || el.flipH || el.flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((el.rotation * Math.PI) / 180);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    // Local copy with the camera and 2-D placement neutralised: rendered at the
    // origin with no scene3d so the recursive call paints a flat body+text, then
    // we warp it. Text selection is dropped (onTextRun omitted) per the note.
    const localEl: ShapeElement = {
      ...el,
      x: 0,
      y: 0,
      rotation: 0,
      flipH: false,
      flipV: false,
      scene3d: undefined,
    };
    const ok = projectScene3dPaint(
      ctx,
      spScene3d.camera,
      x,
      y,
      w,
      h,
      (octx) => {
        // localEl is at the origin (x=y=0) with the element's own EMU size, so
        // the recursive render fills the (0,0,w,h) offscreen at the same scale.
        // No onTextRun → the projected text is not selectable (see the note).
        renderShape(octx, localEl, scale, themeDefaultColor, slideNumber, rc, undefined);
      },
      {
        bevels,
        extrusion: extrusion ?? undefined,
        // Edge margin so the centre-aligned stroke's outer half and the
        // extrusion sweep aren't clipped by the box-sized offscreen (see
        // Project3dOpts.edgePadCss).
        edgePadCss:
          (el.stroke ? (el.stroke.width * scale) / 2 : 0) +
          (el.sp3d?.contourW ? el.sp3d.contourW * scale : 0) +
          (extrusion ? Math.hypot(extrusion.offsetX, extrusion.offsetY) / ctxDevScale : 0) +
          2,
      },
    );
    if (ok) {
      ctx.restore();
      return;
    }
    // Headless fallback (no offscreen): draw flat below, but still without the
    // projection. Restore and fall through to the normal path.
    ctx.restore();
  }

  ctx.save();
  if (el.rotation !== 0 || el.flipH || el.flipV) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((el.rotation * Math.PI) / 180);
    if (el.flipH) ctx.scale(-1, 1);
    if (el.flipV) ctx.scale(1, -1);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }

  const geom = el.geometry.toLowerCase();
  const fillStyle = resolveShapeFill(el.fill, ctx, x, y, w, h);

  // Apply shadow before fill/stroke drawing; ctx.restore() will clear it.
  // The Canvas API exposes a single shadow slot, so when both an outer shadow
  // and a glow are configured we let the outer shadow win (visually dominant)
  // and fall back to the glow only when no outer shadow is present. This is
  // a common — and conservative — interpretation of layered effectLst.
  applyShadow(ctx, el.shadow ?? null, scale);
  if (!el.shadow) applyGlow(ctx, el.glow ?? null, scale);

  const CONNECTOR_GEOMS = new Set([
    'line', 'straightconnector1',
    'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
    'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
  ]);

  // callout family — `<a:ln><a:headEnd|tailEnd>` decorate the LEADER line
  // (the geometry's last `<path>` in presets.json), not the text rectangle
  // (path 0) or the accent bar. callout1 = 2-point leader, callout2/3 =
  // 3-/4-point polylines; getConnectorAnchors resolves the leader's two ends
  // for every variant, so all twelve share the connector decoration path.
  const CALLOUT_GEOMS = new Set([
    'callout1', 'callout2', 'callout3',
    'bordercallout1', 'bordercallout2', 'bordercallout3',
    'accentcallout1', 'accentcallout2', 'accentcallout3',
    'accentbordercallout1', 'accentbordercallout2', 'accentbordercallout3',
  ]);

  // A leader line we re-stroke retracted from its filled decorations: every
  // callout plus the straight / bent (polyline) connectors. Curved connectors
  // are excluded — their leader is a Bézier and getConnectorAnchors.vertices only
  // captures segment endpoints, so retracting from a vertex would straighten the
  // curve. Those keep their preset-engine leader (decoration overlap unchanged).
  const isRetractableLeader = (g: string): boolean =>
    CALLOUT_GEOMS.has(g) || g === 'line' || g === 'straightconnector1' || g.startsWith('bentconnector');

  // ── Dispatch to preset engine when possible ────────────────────────────
  // Preference order: custGeom → generic preset engine → legacy switch.
  // `arc` (ECMA-376 §20.1.10.56 ST_ShapeType "arc") goes through the engine
  // too: its presetShapeDefinitions geometry is a two-<path> shape — path 0
  // (stroke="false") fills the pie wedge (arc + lnTo centre + close) and path 1
  // (fill="none") strokes only the open arc edge. The engine honours those
  // per-path fill/stroke flags, so arc renders its true pie-wedge fill + open
  // outline. The legacy buildShapePath could only draw the open arc (filling it
  // auto-closes into a *chord*), which is why arc used to be excluded here.
  const usePresetEngine =
    !el.custGeom && hasPreset(geom);

  /**
   * Paint the shape's body (fill + stroke) into an arbitrary target context,
   * dispatching through the same preset / custGeom / legacy paths as the live
   * render. Extracted into a closure so the edge/blur effect helpers
   * (innerShdw §20.1.8.40, softEdge §20.1.8.53, reflection §20.1.8.50) can
   * re-paint the shape onto auxiliary canvases.
   *
   * `silhouette`: when set, paint a flat opaque silhouette in that colour
   * (filled path only, no stroke, no gradients) — used by innerShdw to build
   * its mask. The silhouette honours the same even-odd / preset path topology.
   */
  const paintShapeBody = (
    target: CanvasRenderingContext2D,
    silhouette?: string,
  ): void => {
    const tFill = silhouette ?? fillStyle;
    const tStroke = silhouette
      ? null
      : el.stroke
        ? () => {
            applyStroke(target, el.stroke!, scale);
            target.stroke();
          }
        : null;
    const tClearShadow = () => clearShadow(target);

    if (usePresetEngine && !silhouette) {
      renderPresetShape(
        target, geom, x, y, w, h,
        [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8],
        tFill, tStroke, tClearShadow,
        // A retractable leader (callout / straight / bent connector) is
        // re-stroked retracted from its decorated ends in the line-end block
        // below, so suppress the preset engine's full-length leader stroke to
        // avoid a double line / a cap poking through the arrow tip.
        isRetractableLeader(geom) ? { skipTrailingStroke: true } : undefined,
      );
      return;
    }

    target.beginPath();
    if (el.custGeom && el.custGeom.length > 0) {
      buildCustomPath(target, el.custGeom, x, y, w, h);
    } else if (usePresetEngine) {
      // Silhouette of a preset shape: build a single filled outline. Preset
      // path 0 is the body outline (secondary paths are highlights), so the
      // legacy buildShapePath gives a faithful silhouette of the body.
      buildShapePath(target, geom, x, y, w, h, el.adj, el.adj2, el.adj3, el.adj4);
    } else {
      buildShapePath(target, geom, x, y, w, h, el.adj, el.adj2, el.adj3, el.adj4);
    }
    // Normal arc bodies render through the preset engine above; this legacy
    // path is only reached for an arc as a custGeom/effect silhouette built by
    // buildShapePath, which draws the OPEN arc — filling that auto-closes into
    // a chord, not the pie wedge. So skip the fill for arc here (the engine,
    // not this branch, owns arc's pie-wedge fill).
    if (tFill && geom !== 'arc') {
      target.fillStyle = tFill;
      if (geom === 'donut' || geom === 'smileyface' || geom === 'frame') {
        target.fill('evenodd');
      } else {
        target.fill();
      }
      if (!silhouette) tClearShadow();
    }
    if (tStroke) {
      tStroke();
    }
  };

  // ── effectLst edge/blur effects (independent siblings, ECMA-376 §20.1.8.25)
  // Device-pixel canvas extent for the auxiliary effect canvases.
  const deviceW = (ctx.canvas as { width: number }).width || 0;
  const deviceH = (ctx.canvas as { height: number }).height || 0;
  // The effect helpers operate in DEVICE pixels: the aux silhouette is painted
  // through the live transform (which already folds in devicePixelRatio +
  // rotation + flip), and the blit happens at identity. So bbox / radii passed
  // to the helpers must be in device px too. The uniform device scale equals
  // sqrt(|det|) of the transform's linear part — for scale(dpr)·rot·flip this
  // is exactly dpr, independent of rotation or mirroring.
  const liveTransform = ctx.getTransform();
  const det = Math.abs(liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  const effBBox = { x: x * devScale, y: y * devScale, w: w * devScale, h: h * devScale };
  const effScale = scale * devScale; // EMU → device px
  const applyLiveTransform = (c: CanvasRenderingContext2D) => {
    c.setTransform(liveTransform);
  };

  // Reflection sits BEHIND/below the shape — draw it first so the body paints
  // on top. §20.1.8.50. The aux silhouette bakes in the live rotation/flip via
  // setTransform, and the helper's mirror transform operates in device space,
  // so the live ctx must blit at identity.
  if (el.reflection && deviceW > 0 && deviceH > 0) {
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applyReflection(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D); },
      effBBox, el.reflection, effScale, deviceW, deviceH,
    );
    ctx.restore();
  }

  // softEdge feathers the whole body, so it REPLACES the direct body paint.
  // §20.1.8.53. When absent, paint the body normally.
  if (el.softEdge && deviceW > 0 && deviceH > 0) {
    // The feathered body is composited in untransformed device space, so reset
    // the live transform around the blit and re-apply the same transform when
    // painting the body into the aux canvas.
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applySoftEdge(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D); },
      effBBox, el.softEdge, effScale, deviceW, deviceH,
      // Mask is the flat filled silhouette (no stroke) — see applySoftEdge.
      (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D, '#000'); },
    );
    ctx.restore();
  } else {
    paintShapeBody(ctx);
  }

  // innerShdw casts inward, ON TOP of the fill. §20.1.8.40. Composite after the
  // body. The silhouette callback paints a flat opaque mask.
  if (el.innerShadow && deviceW > 0 && deviceH > 0) {
    ctx.save();
    ctx.setTransform(new DOMMatrix());
    applyInnerShadow(
      ctx, (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintShapeBody(c as CanvasRenderingContext2D, '#000'); },
      effBBox, el.innerShadow, effScale, deviceW, deviceH,
    );
    ctx.restore();
  }

  if (el.stroke && (CONNECTOR_GEOMS.has(geom) || CALLOUT_GEOMS.has(geom))) {
    // Connectors and callouts both decorate a *leader line* whose two ends +
    // outward tangents are resolved by getConnectorAnchors from the geometry's
    // last `<path>` (presets.json). For a connector that is the line itself;
    // for a callout it is the attach→tip leader (callout1 straight, callout2/3
    // polyline). headEnd sits on the attach end, tailEnd on the tip — exactly
    // as the `m … l …` order of the preset's leader path dictates.
    const anchors = getConnectorAnchors(geom, x, y, w, h, [el.adj, el.adj2, el.adj3, el.adj4, el.adj5, el.adj6, el.adj7, el.adj8]);
    if (anchors) {
      // ECMA-376 §20.1.8.42 — compound line styles. For straight lines /
      // connectors we re-stroke the segment with multiple parallel lines
      // along the perpendicular of the line direction. Curved connectors and
      // callout leaders fall through to the single-stroke fast path (parallel
      // curves / polylines are a non-trivial geometric operation).
      const cmpd = el.stroke.cmpd;
      const isStraight = geom === 'line' || geom === 'straightconnector1';
      // Retractable leaders (callout / straight / bent connector): paintShapeBody
      // suppressed the preset leader stroke, so re-stroke the polyline here with
      // each decorated end pulled back by the decoration's length, so the line
      // stops at the arrow base instead of poking through its tip (PowerPoint
      // behaviour). Filled ends (triangle/stealth/diamond/oval) retract; open
      // `arrow` / `none` do not (lineEndRetract → 0). A compound straight segment
      // is drawn by drawCompoundLine below instead, so skip the retract there.
      if (isRetractableLeader(geom) && anchors.vertices.length >= 2 && !(cmpd && isStraight)) {
        const pts = anchors.vertices.map((v) => ({ x: v.x, y: v.y }));
        if (el.stroke.tailEnd) {
          const r = lineEndRetract(el.stroke.tailEnd, el.stroke, scale);
          pts[pts.length - 1] = retractLineEndpoint(pts[pts.length - 1], pts[pts.length - 2], r);
        }
        if (el.stroke.headEnd) {
          const r = lineEndRetract(el.stroke.headEnd, el.stroke, scale);
          pts[0] = retractLineEndpoint(pts[0], pts[1], r);
        }
        applyStroke(ctx, el.stroke, scale);
        ctx.beginPath();
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y);
        ctx.stroke();
      }
      if (cmpd && isStraight) {
        drawCompoundLine(ctx, anchors.start, anchors.end, el.stroke, cmpd, scale);
      }
      if (el.stroke.tailEnd) {
        drawArrowHead(ctx, anchors.end.x, anchors.end.y, anchors.end.angle, el.stroke.tailEnd, el.stroke, scale);
      }
      if (el.stroke.headEnd) {
        drawArrowHead(ctx, anchors.start.x, anchors.start.y, anchors.start.angle, el.stroke.headEnd, el.stroke, scale);
      }
    }
  } else if (
    el.stroke &&
    el.custGeom &&
    el.custGeom.length > 0 &&
    ((el.stroke.headEnd && el.stroke.headEnd.type !== 'none') ||
      (el.stroke.tailEnd && el.stroke.tailEnd.type !== 'none'))
  ) {
    // Freeform / curve (custGeom) lines also carry `<a:ln><a:headEnd|tailEnd>`.
    // The connector/callout branches above only cover preset geometries, so a
    // custGeom path's arrow heads were dropped. Extract the open path's two
    // terminal points + outward tangents and decorate them like a connector.
    // Endpoints on a *closed* sub-path are returned as null (PowerPoint draws no
    // line-end decoration on a closed contour).
    const { start, end } = getCustGeomEndpoints(el.custGeom);
    // The endpoint tangent is expressed in normalised (0..1) space; convert to
    // device space accounting for anisotropic scaling (w ≠ h) before atan2 so
    // the arrow head orientation is correct on non-square boxes.
    if (start && el.stroke.headEnd && el.stroke.headEnd.type !== 'none') {
      const sx = x + start.x * w;
      const sy = y + start.y * h;
      const sAngle = Math.atan2(start.dy * h, start.dx * w);
      drawArrowHead(ctx, sx, sy, sAngle, el.stroke.headEnd, el.stroke, scale);
    }
    if (end && el.stroke.tailEnd && el.stroke.tailEnd.type !== 'none') {
      const ex = x + end.x * w;
      const ey = y + end.y * h;
      const eAngle = Math.atan2(end.dy * h, end.dx * w);
      drawArrowHead(ctx, ex, ey, eAngle, el.stroke.tailEnd, el.stroke, scale);
    }
  }

  // Render text inside the rotation context so text follows shape rotation
  if (el.textBody) {
    const defaultTextColor = shapeDefaultTextColor(el, rc);
    ctx.save();
    if (el.flipH || el.flipV) {
      const cx = x + w / 2;
      const cy = y + h / 2;
      // The shape itself stays mirrored, but text should remain readable.
      // Apply the same flip again around the shape centre to cancel only the text mirror.
      ctx.translate(cx, cy);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-cx, -cy);
    }
    // For ellipses, PowerPoint positions text relative to the inscribed rectangle
    // (the maximum-area rectangle that fits inside the ellipse: sides = a/√2, b/√2).
    // This only affects non-ctr anchors; ctr anchor is invariant to this inset.
    let tx = x, ty = y, tw = w, th = h;
    if (el.textRect) {
      // SmartArt drawings carry an explicit text frame (<dsp:txXfrm>) that
      // PowerPoint pre-computed against the actual layout — e.g. an arrow's
      // label sits past an overlapping circle node, a roundRect's label avoids
      // an overlapping bottom badge. Honour it verbatim (insets apply within).
      tx = emuToPx(el.textRect.x, scale);
      ty = emuToPx(el.textRect.y, scale);
      tw = emuToPx(el.textRect.width, scale);
      th = emuToPx(el.textRect.height, scale);
    } else if (geom === 'ellipse') {
      const insetX = w * (1 - 1 / Math.SQRT2) / 2;
      const insetY = h * (1 - 1 / Math.SQRT2) / 2;
      tx = x + insetX; ty = y + insetY;
      tw = w / Math.SQRT2; th = h / Math.SQRT2;
    } else {
      // Preset text rectangle (ECMA-376 §20.1.9.21): e.g. an arrow's label sits
      // in the shaft, not the full bbox. Insets (lIns/…) apply within this rect.
      const tr = presetTextRect(geom, x, y, w, h, el.adj, el.adj2);
      if (tr) { tx = tr.tx; ty = tr.ty; tw = tr.tw; th = tr.th; }
    }
    // Pass el.rotation so the text-layer overlay can CSS-rotate the shape div to match.
    renderTextBody(ctx, el.textBody, tx, ty, tw, th, scale, defaultTextColor, el.rotation, false, false, themeDefaultColor, slideNumber, rc, shapeTextRunCallback, false, fetchImage);
    ctx.restore();
  }

  ctx.restore();
}

/**
 * Build a canvas path from custGeom path commands.
 * Coordinates are in [0,1] relative to the shape bounding box;
 * the renderer maps them to canvas pixels.
 * Tracks pen position so arcTo can compute the ellipse centre correctly.
 */
const buildCustomPath = buildCustomPathCore;

/**
 * Format an autoNum bullet label from a counter value and OOXML numType.
 * Spec: ECMA-376 §20.1.10.61 (ST_TextAutonumberScheme).
 *
 * Supports the symmetric Plain / Period / ParenR / ParenBoth variants for
 * the four core scripts (arabic, alphaLc/Uc, romanLc/Uc) plus arabicDb
 * (full-width Arabic digits). CJK / Thai / Hebrew / Arabic-Abjad schemes
 * are intentionally not handled here — they fall through to the default
 * `arabicPeriod` rendering rather than emitting a wrong-looking glyph.
 */
function formatAutoNum(counter: number, numType: string): string {
  const arabic = `${counter}`;
  const alphaLc = counter >= 1 && counter <= 26
    ? String.fromCharCode(96 + counter)
    : arabic;
  const alphaUc = counter >= 1 && counter <= 26
    ? String.fromCharCode(64 + counter)
    : arabic;
  const romanLc = toRoman(counter).toLowerCase();
  const romanUc = toRoman(counter);
  // ECMA-376 §20.1.10.61 lists `arabicDb*` as full-width Arabic digits
  // (U+FF10–U+FF19), used in East Asian numbered lists.
  const arabicDb = arabic.replace(/[0-9]/g, (d) =>
    String.fromCharCode(0xff10 + (d.charCodeAt(0) - 0x30)),
  );

  switch (numType) {
    case 'arabicPlain':       return arabic;
    case 'arabicPeriod':      return `${arabic}.`;
    case 'arabicParenR':      return `${arabic})`;
    case 'arabicParenBoth':   return `(${arabic})`;
    case 'arabicDbPlain':     return arabicDb;
    case 'arabicDbPeriod':    return `${arabicDb}.`;
    case 'alphaLcPlain':      return alphaLc;
    case 'alphaLcPeriod':     return `${alphaLc}.`;
    case 'alphaLcParenR':     return `${alphaLc})`;
    case 'alphaLcParenBoth':  return `(${alphaLc})`;
    case 'alphaUcPlain':      return alphaUc;
    case 'alphaUcPeriod':     return `${alphaUc}.`;
    case 'alphaUcParenR':     return `${alphaUc})`;
    case 'alphaUcParenBoth':  return `(${alphaUc})`;
    case 'romanLcPlain':      return romanLc;
    case 'romanLcPeriod':     return `${romanLc}.`;
    case 'romanLcParenR':     return `${romanLc})`;
    case 'romanLcParenBoth':  return `(${romanLc})`;
    case 'romanUcPlain':      return romanUc;
    case 'romanUcPeriod':     return `${romanUc}.`;
    case 'romanUcParenR':     return `${romanUc})`;
    case 'romanUcParenBoth':  return `(${romanUc})`;
    default:                  return `${arabic}.`;
  }
}

function toRoman(n: number): string {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}

/**
 * True when a paragraph has renderable content — at least one non-empty text
 * run, or any equation run. Line breaks alone (a paragraph the user created by
 * pressing Enter on an empty line) do NOT count as content.
 *
 * PowerPoint does not draw a bullet/number marker on an empty paragraph and
 * does not advance an autoNum counter for it (a blank line between "1." and
 * "2." stays unnumbered and the sequence continues). ECMA-376 §21.1.2.4.x
 * (CT_Bullet / CT_TextCharBullet / CT_TextAutoNumberBullet) gives no numeric
 * rule for this, so PowerPoint's behaviour is the reference: suppress the
 * marker whenever the paragraph carries no real content. This is the plain
 * "is there content?" test, not a sample-specific heuristic.
 */
export function paragraphHasRenderableContent(para: Paragraph): boolean {
  for (const r of para.runs) {
    if (r.type === 'text' && r.text !== '') return true;
    if (r.type === 'math') return true;
  }
  return false;
}

/**
 * Resolve the bullet/number marker *label string* for a paragraph, mutating
 * the per-level autoNum counter map as a side effect. Returns '' (= no marker)
 * for `none`/`inherit` bullets and for empty paragraphs.
 *
 * Empty-paragraph handling (PowerPoint reference behaviour, see
 * `paragraphHasRenderableContent`): when the paragraph has no renderable
 * content we draw no marker AND do not advance the autoNum counter, so a blank
 * line between numbered items keeps the sequence going (… "1." / "" / "2." …).
 *
 * The counter-reset semantics (clear on char bullets / non-list paragraphs)
 * are intentionally kept here so the same map walk the renderer relied on is
 * exercised by the unit tests. Font and colour resolution stay in the renderer
 * because they depend on canvas/theme context.
 */
export function resolveBulletLabel(
  para: Paragraph,
  autoNumCounters: Map<number, number>,
): string {
  const hasContent = paragraphHasRenderableContent(para);
  if (para.bullet.type === 'char') {
    // Reset counters when switching to char bullets.
    autoNumCounters.clear();
    return hasContent ? symbolFontToUnicode(para.bullet.char, para.bullet.fontFamily ?? null) : '';
  }
  if (para.bullet.type === 'autoNum') {
    if (!hasContent) return '';
    const lvl = para.lvl;
    if (!autoNumCounters.has(lvl)) {
      autoNumCounters.set(lvl, para.bullet.startAt ?? 1);
    } else {
      autoNumCounters.set(lvl, autoNumCounters.get(lvl)! + 1);
    }
    return formatAutoNum(autoNumCounters.get(lvl)!, para.bullet.numType);
  }
  // none / inherit — not a list paragraph; reset autoNum counters.
  autoNumCounters.clear();
  return '';
}

// Exported (like `layoutParagraph` / `paintHighlight`) so the picture-bullet
// draw path can be unit-tested against a mock 2D context without standing up a
// full canvas. Not re-exported from index.ts — module-internal otherwise.
export function renderTextBody(
  ctx: CanvasRenderingContext2D,
  body: TextBody,
  bx: number,
  by: number,
  bw: number,
  bh: number,
  scale: number,
  shapeDefaultTextColor: string | null = null,
  shapeRotation = 0,
  shapeFlipH = false,
  shapeFlipV = false,
  themeDefaultColor = '#000000',
  slideNumber?: number,
  rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
  onTextRun?: TextRunCallback,
  measureOnly = false,
  fetchImage?: FetchImage,
  // ECMA-376 §20.1.10.83 `eaVert` ("East Asian Vertical"): CJK glyphs stand
  // UPRIGHT while Latin rotates. Set true ONLY when re-entering horizontally
  // inside the +90°-rotated `eaVert` frame (below), so the run-draw loop applies
  // per-glyph UAX#50 orientation via `drawEaVertRun`. False for `horz` and for
  // `vert`/`vert270` (whose spec meaning IS to rotate every glyph), keeping those
  // paths byte-identical.
  eaVertUpright = false,
): number | void {
  // Vertical text: rotate rendering context so text flows top-to-bottom.
  // "vert" and "eaVert" both approximate to 90° clockwise rotation.
  // "vert270" rotates 270° (= 90° counterclockwise).
  const isVert    = body.vert === 'vert' || body.vert === 'eaVert';
  const isVert270 = body.vert === 'vert270';

  if (isVert || isVert270) {
    // Set up a rotated coordinate space:
    // Centre of the bounding box remains fixed; swap w and h for the text layout.
    const cx = bx + bw / 2;
    const cy = by + bh / 2;
    const vertRot = isVert ? 90 : -90;

    // Wrap onTextRun to convert from the rotated sub-frame back to the original
    // shape frame so that _buildTextLayer can apply a single CSS rotation.
    //
    // In the recursive call the origin is (-bh/2, -bw/2) with axes (bh, bw).
    // For a run at canvas (penX, cursorY) in that sub-frame:
    //   inShapeX_rec = penX + bh/2,  inShapeY_rec = cursorY + bw/2
    //
    // We need the position in the *original* shape frame so that after
    // CSS rotate(shapeRotation + vertRot) the span lands on the same pixel:
    //   inShapeX_span = penX + bw/2 = inShapeX_rec - bh/2 + bw/2
    //   inShapeY_span = cursorY + bh/2 = inShapeY_rec - bw/2 + bh/2
    const wrappedOnTextRun: TextRunCallback | undefined = onTextRun
      ? (run) => onTextRun({
          ...run,
          inShapeX: run.inShapeX - bh / 2 + bw / 2,
          inShapeY: run.inShapeY - bw / 2 + bh / 2,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
          textBodyRotation: vertRot,
        })
      : undefined;

    if (measureOnly) {
      // The rotated sub-frame's content height runs along the original width
      // axis; for a table row the vertical extent is the original box width.
      return bw;
    }
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(isVert270 ? -Math.PI / 2 : Math.PI / 2);
    // After rotation the "width" direction of the new frame is the original height.
    // `eaVert` re-enters with per-glyph uprighting (UAX#50): CJK stands up, Latin
    // stays sideways. `vert`/`vert270` re-enter WITHOUT it — their spec meaning is
    // to rotate every glyph — so those stay byte-identical to before.
    renderTextBody(ctx, { ...body, vert: 'horz' }, -bh / 2, -bw / 2, bh, bw, scale, shapeDefaultTextColor, 0, false, false, themeDefaultColor, slideNumber, rc, wrappedOnTextRun, false, fetchImage, body.vert === 'eaVert');
    ctx.restore();
    return;
  }

  // ── WordArt text warp (ECMA-376 §20.1.9.19, prstTxWarp) ──────────────────
  // When the body carries a known warp preset, each glyph is mapped through the
  // preset's envelope rather than laid out flat. This is a *separate* draw path
  // gated on the warp's presence, so every unwarped body renders exactly as
  // before (byte-identical). measureOnly bodies fall through to the flat pass —
  // a warped shape is fixed-size, not a table cell auto-sizing to its text.
  const warp = body.textWarp;
  if (!measureOnly && warp && hasTextWarp(warp.preset)) {
    renderWarpedText(
      ctx,
      body,
      warp.preset,
      warp.adj ?? [],
      bx,
      by,
      bw,
      bh,
      scale,
      shapeDefaultTextColor ?? themeDefaultColor,
      rc,
    );
    return;
  }

  const lPad = emuToPx(body.lIns, scale);
  const rPad = emuToPx(body.rIns, scale);
  const tPad = emuToPx(body.tIns, scale);
  const bPad = emuToPx(body.bIns, scale);
  // ECMA-376 §20.1.10.5 spAutoFit: "the shape's height and width should resize
  // to accommodate the text". The bbox is no longer a wrap boundary in that
  // mode — the shape grows to fit. Treating spAutoFit as wrap=none on the
  // horizontal axis keeps mixed-size sequences (e.g. "20代", "YoY+11.9%")
  // on one line, matching PowerPoint. Vertical growth is handled below by
  // the spAutoFit branch in the height calculation.
  // ECMA-376 §20.1.10.5 spAutoFit + §20.1.10.7 wrap interaction:
  // - `wrap="none"`: text never wraps (regardless of autoFit).
  // - `wrap="square"` + `<a:spAutoFit/>`: PowerPoint auto-fits the SHAPE to
  //   the text. If the text's natural single-line width fits the bbox, the
  //   shape stays at its current width and the text doesn't wrap (sample-2
  //   slide-13's "20代" textbox: ~50px text in ~70px bbox → no wrap). If
  //   the natural width exceeds the bbox, text wraps and the shape grows
  //   vertically (sample-2 slide-16's "1Q業績 要因" callouts: a single
  //   long Japanese paragraph that has to wrap within a fixed-width bbox).
  // - `wrap="square"` (default) without spAutoFit: always wrap.
  //
  // The pre-pass below measures the paragraphs' total natural width; if
  // every paragraph fits the bbox without wrapping, we keep the spAutoFit
  // "no-wrap" semantics. Otherwise we wrap normally so the long text
  // doesn't run off the side of the shape.
  const baseDoWrap = body.wrap !== 'none';
  const isSpAutoFit = body.autoFit === 'sp';
  const doWrap = isSpAutoFit
    ? (baseDoWrap && naturalWidthExceedsBbox(ctx, body, bw, lPad, rPad, scale, rc))
    : baseDoWrap;
  // ECMA-376 §20.1.10.34 numCol — distribute paragraphs across N text columns.
  const numCol = Math.max(1, body.numCol ?? 1);
  const spcColPx = emuToPx(body.spcCol ?? 0, scale);

  const bodyDefaultBold   = body.defaultBold   ?? false;
  const bodyDefaultItalic = body.defaultItalic ?? false;
  const bodyDefaultColor = shapeDefaultTextColor ?? themeDefaultColor;

  // ── Pass 1: lay out all paragraphs ──────────────────────────────────────

  interface LineEntry {
    line: LayoutLine;
    linePx: number;       // spacing advancement (lineHeight + spaceAfter for last line)
    lineHeight: number;   // pure line height used for baseline positioning (without spaceAfter)
    topGapPx: number;     // spaceBefore for first line of paragraph
    textXOffset: number;  // additional X offset for first-line indent (non-bullet)
    bulletLabel: string;  // text to render as bullet ('' = none)
    bulletFont: string;
    bulletColor: string;
    bulletX: number;      // canvas X for bullet
    // Picture bullet (`<a:buBlip>`, §21.1.2.4.2): the resolved image + its
    // drawn size in px (square, scaled by buSzPct). null when this paragraph
    // has no picture bullet. Only set on the paragraph's first line.
    bulletImage: { imagePath: string; mimeType: string; sizePx: number } | null;
    textX: number;        // canvas X for text
    textMaxW: number;     // max wrap width
    alignment: string;
    isLastLine: boolean;
    para: Paragraph;
  }

  // buildLayout runs Pass 1 at a given font scale (1.0 = normal; <1 = normAutoFit shrink)
  const buildLayout = (fontScale: number): { allLines: LineEntry[], totalHeight: number } => {
  const bodyDefaultFontSizePx = (body.defaultFontSize ?? 18) * PT_TO_EMU * scale * fontScale;
  const allLines: LineEntry[] = [];
  let totalHeight = 0;

  // AutoNum counters per list level
  const autoNumCounters = new Map<number, number>();

  for (let paraIdx = 0; paraIdx < body.paragraphs.length; paraIdx++) {
    const para = body.paragraphs[paraIdx];
    const marLPx   = emuToPx(para.marL,   scale);
    const marRPx   = emuToPx(para.marR,   scale);
    const indentPx = emuToPx(para.indent, scale);

    // Para-level defaults (cascade: para defRPr → body default)
    const paraDefaultFontSizePx = para.defFontSize != null
      ? para.defFontSize * PT_TO_EMU * scale * fontScale : bodyDefaultFontSizePx;
    const paraDefaultColor = para.defColor
      ? hexToRgba(para.defColor) : bodyDefaultColor;

    // Bullet resolution. A picture bullet (`blip`) occupies the gutter exactly
    // like a char/autoNum marker, so it must suppress the first-line hanging
    // indent too — otherwise the first line of a hanging-indent list starts at
    // the bullet's x and renders ON TOP of the picture (cf. the char-bullet em-
    // dash overlap noted below, and docx PR #476).
    const hasBullet = paragraphHasBullet(para);

    // Per ECMA-376 §21.1.2.4.13: when no buSz* is declared, the bullet takes
    // the first run's font size. Using paraDefaultFontSizePx here (the layout
    // lvl1pPr defRPr fallback, typically 18pt) oversizes the bullet so a
    // hanging indent calibrated against the run (12pt) can't contain it —
    // that's why the em-dash was overlapping the text.
    const firstRunSizePt = (() => {
      for (const r of para.runs) {
        if (r.type === 'text' && r.fontSize != null) return r.fontSize;
      }
      return null;
    })();
    const bulletBaseSizePx = firstRunSizePt != null
      ? firstRunSizePt * PT_TO_EMU * scale * fontScale
      : paraDefaultFontSizePx;

    // ECMA-376 §21.1.2.4.4 (CT_TextCharBullet) / §21.1.2.4.10 (buClrTx): when
    // no explicit `<a:buClr>` is present, the bullet inherits the *first run*'s
    // color, NOT the shape-level default text color. The two diverge on
    // templates where `<p:style><a:fontRef>` resolves to white (lt1) — runs
    // override that with their own `<a:rPr><a:solidFill>`, but bullets without
    // a buClr would otherwise pick up the white default and become invisible
    // (the slide-13 sample-2 regression).
    const firstRunColorHex = (() => {
      for (const r of para.runs) {
        if (r.type === 'text' && r.color) return r.color;
      }
      return null;
    })();
    const bulletInheritedColor = firstRunColorHex
      ? hexToRgba(firstRunColorHex)
      : paraDefaultColor;

    let bulletLabel  = '';
    let bulletFont   = buildFont(false, false, bulletBaseSizePx, 'sans-serif', rc);
    let bulletColor  = bulletInheritedColor;
    // Picture bullet (`<a:buBlip>`, §21.1.2.4.2). Resolved to its image + drawn
    // size below; stays null for char/number/none bullets.
    let bulletImage: { imagePath: string; mimeType: string; sizePx: number } | null = null;

    // Resolve the marker label and advance the autoNum counter. Empty
    // paragraphs (Enter on a blank line) draw no marker and do not advance the
    // counter — PowerPoint keeps a blank line between numbered items unnumbered
    // while continuing the sequence. ECMA-376 §21.1.2.4.x gives no numeric
    // rule, so PowerPoint's behaviour is the reference (see resolveBulletLabel).
    bulletLabel = resolveBulletLabel(para, autoNumCounters);

    // The parser may emit the picture-bullet variant (`type: "blip"`), which the
    // shared core `Bullet` type doesn't list — narrow once via asBullet.
    const bullet = asBullet(para.bullet);
    if (bullet.type === 'char') {
      const b = bullet;
      // ECMA-376 §21.1.2.4.10 (buSzPts): an ABSOLUTE point size sizes the marker
      // independent of the run — convert pt → px like any run size
      // (PT_TO_EMU·scale·fontScale). It wins over §21.1.2.4.9 (buSzPct, percent of
      // the run size); with neither, the marker takes the run size (§21.1.2.4.13).
      const bSizePx = b.sizePts != null
        ? b.sizePts * PT_TO_EMU * scale * fontScale
        : b.sizePct != null
          ? bulletBaseSizePx * (b.sizePct / 100)
          : bulletBaseSizePx;
      // If the char was mapped to a Unicode symbol, use sans-serif for reliable rendering.
      // Otherwise use the specified font (e.g. Wingdings on systems that have it).
      const convertedFamily = bulletLabel !== b.char ? 'sans-serif' : normalizeFontFamily(b.fontFamily ?? null, rc);
      bulletFont  = buildFont(false, false, bSizePx, convertedFamily, rc);
      bulletColor = b.color ? hexToRgba(b.color) : bulletInheritedColor;
    } else if (bullet.type === 'autoNum') {
      bulletFont  = buildFont(false, false, bulletBaseSizePx, 'sans-serif', rc);
      // ECMA-376 §21.1.2.4.4 (buClr): an explicit `<a:buClr>` colours the
      // auto-number marker, mirroring the char-bullet branch above. Only when it
      // is absent does the marker fall back to the buClrTx default
      // (§21.1.2.4.5 — the inherited first-run colour).
      bulletColor = bullet.color ? hexToRgba(bullet.color) : bulletInheritedColor;
    } else if (bullet.type === 'blip') {
      // ECMA-376 §21.1.2.4.2 picture bullet. The bitmap is drawn as a square
      // sized to the text (the bullet's em box), scaled by `<a:buSzPct>`
      // (§21.1.2.4.9; default 100%). It's not a glyph, so there is no label —
      // an empty paragraph still draws no marker (bulletLabel stays '' and the
      // draw site gates the image on the first line having content).
      // §21.1.2.4.10 (buSzPts): an absolute point size overrides the §21.1.2.4.9
      // percentage, matching the char-bullet branch above.
      const b = bullet;
      const sizePx = b.sizePts != null
        ? b.sizePts * PT_TO_EMU * scale * fontScale
        : b.sizePct != null
          ? bulletBaseSizePx * (b.sizePct / 100)
          : bulletBaseSizePx;
      bulletImage = { imagePath: b.imagePath, mimeType: b.mimeType, sizePx };
    }

    // Text start X and wrap width.
    //
    // ECMA-376 §20.1.10.34 numCol — the bbox is split into N equal columns
    // separated by `spcCol` gutters, and paragraphs flow column-by-column.
    // For Pass 1 we lay out at the column width so wrapping & line counts are
    // correct; Pass 2 reuses the same lines and just shifts the X origin per
    // column. textX/bulletX below are relative to column 0; the column shift
    // is added later when we walk `allLines`.
    const colWidth = numCol > 1
      ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol
      : bw - lPad - rPad;
    const textX    = bx + lPad + marLPx;
    // The marker seats at the RAW (signed) indent — a hanging gutter is negative,
    // so this is deliberately NOT routed through firstLineIndentPxFor (which is
    // the first-line TEXT shift, clamped ≥ 0). Keep them distinct.
    const bulletX  = bx + lPad + marLPx + indentPx;
    const textMaxW = colWidth - marLPx - marRPx;

    const maxW = doWrap ? textMaxW : Infinity;
    // A positive first-line indent narrows ONLY the first line's wrap budget
    // (continuation lines keep the full width). firstLineIndentPxFor keeps this
    // in lockstep with the draw-side `textXOffset` (§below) and the
    // naturalWidthExceedsBbox measurement; a bullet's gutter / a negative
    // (hanging) indent contribute 0.
    const firstLineIndentPx = firstLineIndentPxFor(hasBullet, indentPx);
    const lines = layoutParagraph(ctx, para, maxW, paraDefaultFontSizePx, paraDefaultColor, scale, marLPx, bodyDefaultBold, bodyDefaultItalic, fontScale, slideNumber, rc, firstLineIndentPx);

    // spaceBefore/After are in hundredths of a point → convert to canvas px
    const spaceBeforePx = para.spaceBefore != null ? (para.spaceBefore / 100) * PT_TO_EMU * scale * fontScale : 0;
    const spaceAfterPx  = para.spaceAfter  != null ? (para.spaceAfter  / 100) * PT_TO_EMU * scale * fontScale : 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const isFirst = i === 0;
      const isLast  = i === lines.length - 1;

      // Line height: use the max font size among rendered segments. The layout
      // default (paraDefaultFontSizePx) is used only as a fallback for empty
      // paragraphs — otherwise PowerPoint slide-layouts with placeholder
      // defaults like `defRPr sz="30000"` (300pt prompt-text marker) would
      // inflate lineHeight and push real 24pt runs far below the anchor.
      let maxSizePx = 0;
      // Design single-line-height FLOOR (ECMA-376 §17.3.1.33, shared with docx
      // via core's `intendedSingleLinePx`). PowerPoint sizes single spacing as a
      // flat 1.2×em; for a SUBSTITUTED face whose Windows design line height is
      // taller than that (Meiryo 1.596×em, Sakkal Majalla 1.3965×em) the flat
      // ratio understates Word/PowerPoint's line box and lines overlap. This is
      // a FLOOR, not a replacement: `intendedSingleLinePx` returns 0 for every
      // non-tabled family, so `max(1.2×em, 0)` leaves all other fonts (all Latin,
      // installed CJK, etc.) exactly on PowerPoint's 1.2×em convention.
      let designSingle = 0;
      for (const seg of line.segments) {
        // For an equation, the line must be at least as tall as its own font
        // size (so a short label like "y"/"p"/"z" gets the normal font-ascent
        // baseline and sits where PowerPoint puts it — not floated up by its
        // tight ink box), AND tall enough for its visual box (fractions, norms,
        // big operators) to fit inside the 1.2× leading.
        const effSize = seg.math
          ? Math.max(seg.sizePx, (seg.math.ascent + seg.math.descent) / 1.2)
          : seg.sizePx;
        if (effSize > maxSizePx) maxSizePx = effSize;
        // Equations carry no text family; only text segments contribute a floor.
        if (!seg.math) {
          const ds = intendedSingleLinePx(seg.fontFamily, seg.sizePx);
          if (ds > designSingle) designSingle = ds;
        }
      }
      if (maxSizePx === 0) maxSizePx = paraDefaultFontSizePx;
      // Bullet font size also counts
      if (isFirst && bulletLabel) {
        ctx.font = bulletFont;
        const bm = ctx.measureText('M');
        const bSizeApprox = bm.actualBoundingBoxAscent + bm.actualBoundingBoxDescent;
        if (bSizeApprox > maxSizePx) maxSizePx = bSizeApprox;
      }
      // A picture bullet's box also counts toward the line height so a tall
      // bitmap marker isn't clipped by a short first line.
      if (isFirst && bulletImage && bulletImage.sizePx > maxSizePx) {
        maxSizePx = bulletImage.sizePx;
      }

      // Single-line base with the design-line-height floor applied (see the
      // `designSingle` loop above). `singleLine` replaces the bare `maxSizePx *
      // 1.2` everywhere the base is a MULTIPLE of the single line (the pct and
      // no-spaceLine cases); the exact-pt `spcPts` case is an absolute height
      // and is deliberately NOT floored.
      const singleLine = Math.max(maxSizePx * 1.2, designSingle);
      let lineHeight: number;
      if (para.spaceLine) {
        if (para.spaceLine.type === 'pct') {
          // spcPct 100% = single line spacing = natural font leading ≈ 1.2× em
          lineHeight = singleLine * (para.spaceLine.val / 100000);
        } else {
          lineHeight = para.spaceLine.val * PT_TO_EMU * scale;
        }
      } else {
        lineHeight = singleLine;
      }
      // normAutofit lnSpcReduction (ECMA-376 §21.1.2.1.3): PowerPoint reduces
      // each paragraph's line spacing by this fraction alongside the font
      // shrink. Apply it only when normAutofit stored a value AND the paragraph
      // has PERCENTAGE line spacing — the spec's normative note reads "This
      // attribute applies only to paragraphs with percentage line spacing." pct
      // and the implicit single (= 100 % percentage) qualify; an absolute
      // spcPts height does not.
      if (body.autoFit === 'norm' && body.lnSpcReduction != null && para.spaceLine?.type !== 'pts') {
        lineHeight *= 1 - body.lnSpcReduction;
      }
      const linePx  = lineHeight + (isLast ? spaceAfterPx : 0);
      // ECMA-376 §21.1.2.2.6 (a:spcBef): paragraph "space before" is the gap
      // *between* paragraphs. PowerPoint suppresses it on the first paragraph
      // of a text body — otherwise placeholders whose layout-default `spcBef`
      // is 10 pt (sample-1 slide-5 "Figure 1." caption inherits this from the
      // layout body lstStyle) get pushed ~10 px below the placeholder top and
      // collide with the chart title sitting just below in the slide.
      const topGap  = isFirst && paraIdx > 0 ? spaceBeforePx : 0;
      // Non-bullet first-line indent, clamped ≥ 0 via firstLineIndentPxFor so the
      // draw offset matches the wrap budget and the spAutoFit measurement. A
      // negative ("hanging") indent is NOT honored at draw time — it would shift
      // the first line left of marL while the wrap pass keeps full width, so the
      // two would disagree; we clamp both to 0.
      const textXOffset = isFirst ? firstLineIndentPxFor(hasBullet, indentPx) : 0;

      // Picture bullets, like char/number markers, are drawn only on the
      // paragraph's first line and only when that line carries content (an
      // empty paragraph gets no marker — PowerPoint behaviour, mirroring
      // resolveBulletLabel's empty-paragraph handling).
      const lineHasContent = line.segments.some(
        (s) => (s.text && s.text.length > 0) || s.math != null,
      );
      const entryBulletImage = isFirst && lineHasContent ? bulletImage : null;

      allLines.push({
        line, linePx, lineHeight, topGapPx: topGap,
        textXOffset,
        bulletLabel: isFirst ? bulletLabel : '',
        bulletFont, bulletColor, bulletX,
        bulletImage: entryBulletImage,
        textX, textMaxW,
        alignment: para.alignment,
        isLastLine: isLast,
        para,
      });
      totalHeight += linePx + topGap;
    }
  }

  return { allLines, totalHeight };
  }; // end buildLayout

  let { allLines, totalHeight } = buildLayout(1.0);

  // ── normAutoFit ──────────────────────────────────────────────────────────
  // PowerPoint stores the font-shrink ratio it computed at edit time in
  // <a:normAutofit fontScale> (ECMA-376 §21.1.2.1.3). When present, apply it
  // directly — this reproduces PowerPoint's exact layout instead of guessing a
  // scale from our own (slightly different) text metrics. Only when no scale
  // was stored do we fall back to fitting the text by search.
  if (body.autoFit === 'norm') {
    if (body.fontScale != null && body.fontScale > 0) {
      if (body.fontScale < 1.0) ({ allLines, totalHeight } = buildLayout(body.fontScale));
    } else {
      const maxContentH = bh - tPad - bPad;
      if (totalHeight > maxContentH && maxContentH > 0) {
        let lo = 0.1, hi = 1.0;
        for (let i = 0; i < 6; i++) {
          const mid = (lo + hi) / 2;
          if (buildLayout(mid).totalHeight <= maxContentH) lo = mid; else hi = mid;
        }
        ({ allLines, totalHeight } = buildLayout(lo));
      }
    }
  }

  // ── measure-only: return the content height the text body needs ─────────
  // Used by renderTable to grow rows to fit their tallest cell (ECMA-376
  // §21.1.3.18: a:tr@h is a minimum). Returns padding + laid-out text height.
  if (measureOnly) {
    return tPad + totalHeight + bPad;
  }

  // ── anchor="b" with bh=0: auto-height growing upward from by ────────────
  // When cy=0 and anchor="b", off_y is the bottom anchor; shape grows upward.
  const anchor = body.verticalAnchor ?? 't';
  let effectiveBy = by;
  let effectiveBh: number;
  if (bh === 0 && anchor === 'b') {
    effectiveBh = tPad + totalHeight + bPad;
    effectiveBy = by - effectiveBh;
  } else {
    // ── Effective height (spAutoFit: shape expands to fit text) ─────────────
    const isSpAutoFit = body.autoFit === 'sp';
    effectiveBh = isSpAutoFit
      ? Math.max(bh, tPad + totalHeight + bPad)
      : bh;
  }

  // ── Vertical anchor ─────────────────────────────────────────────────────
  let cursorY: number;
  const contentH = Math.max(0, effectiveBh - tPad - bPad);
  if (anchor === 'ctr') {
    cursorY = effectiveBy + tPad + (contentH - totalHeight) / 2;
  } else if (anchor === 'b') {
    cursorY = effectiveBy + effectiveBh - bPad - totalHeight;
  } else {
    cursorY = effectiveBy + tPad;
  }

  // ── Pass 2: render ───────────────────────────────────────────────────────
  ctx.save();
  // penX / baseline are computed manually below, so the canvas text origin
  // must be normalized before fillText() or alignment/anchor math will drift.
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  // ECMA-376 §20.1.2.3.6 (a:bodyPr): PowerPoint does NOT clip text that
  // overflows its shape — the text simply renders past the shape bounds and
  // can overlap with adjacent elements. Our previous behavior clipped, which
  // cropped long text in fixed-height boxes. Only clip when the caller has
  // opted into wrap=none AND a finite x-axis rectangle (rare), which we
  // approximate here by skipping clipping entirely for the default bodyPr.
  // `body.wrap === "none"` means horizontal non-wrap; it doesn't affect
  // clipping per spec either, so we just don't clip.

  // Multi-column flow state. PowerPoint's actual behaviour with `numCol`:
  //
  // - If the total content height fits within a single column, ALL paragraphs
  //   stay in column 0 even when `numCol > 1`. (Sample-2 slide-13's "従来"
  //   box has 4 paragraphs in a `numCol="2"` body; PowerPoint stacks them
  //   vertically because they fit, not 2+2.)
  // - Otherwise, paragraphs are distributed *balanced* (each column gets
  //   roughly ⌈N/numCol⌉ entries) so parallel rows across columns line up.
  //   (Sample-2 slide-13's "新機能" box has 9 paragraphs that overflow a
  //   single column; the right column starts with a blank paragraph so the
  //   blue 利用履歴/価値観/属性詳細 sit at the same y as 年齢/性別/居住地.)
  //
  // A pure overflow-driven advance would either pack column 0 maximally and
  // shift the right column up (breaking row alignment) or never collapse
  // multi-col to single when content fits.
  //
  // We pre-compute, for each entry, which column it belongs to. Column 0
  // starts at the initial cursorY; each subsequent column resets the cursor
  // and shifts X by colWidth + spcCol.
  const colTopY = cursorY;
  const colWidthShift = numCol > 1
    ? (bw - lPad - rPad - (numCol - 1) * spcColPx) / numCol + spcColPx
    : 0;
  const colHeightCapacity = Math.max(0, effectiveBh - tPad - bPad);
  // When `bh === 0` (caller relies on auto-height) the capacity is unbounded
  // and content always fits, so a multi-col directive collapses to single col.
  const fitsInOneCol = bh === 0 || totalHeight <= colHeightCapacity + 0.5;
  const useMultiCol = numCol > 1 && !fitsInOneCol;
  const linesPerCol = useMultiCol ? Math.ceil(allLines.length / numCol) : allLines.length;
  let colIdx = 0;
  let entriesInCol = 0;

  for (const entry of allLines) {
    const { line, linePx, lineHeight, topGapPx, textXOffset, bulletLabel, bulletFont, bulletColor, bulletImage, alignment, isLastLine } = entry;
    // Balanced column advance: when the current column has reached its share
    // of paragraphs, jump to the next one. PowerPoint never breaks a single
    // line across columns and never spills past the last column — anything
    // beyond the last column simply runs past the bbox, matching PPT.
    if (numCol > 1 && colIdx < numCol - 1 && entriesInCol >= linesPerCol) {
      colIdx++;
      entriesInCol = 0;
      cursorY = colTopY;
    }
    cursorY += topGapPx;
    entriesInCol++;
    // §21.1.2.1.1 bodyPr@rtlCol: columns fill right-to-left — mirror the
    // visual column index. LTR bodies (rtlCol false/absent) are unchanged.
    const visCol = body.rtlCol ? numCol - 1 - colIdx : colIdx;
    const xShift = visCol * colWidthShift;
    const textX = entry.textX + xShift;
    const bulletX = entry.bulletX + xShift;
    const textMaxW = entry.textMaxW;

    // Bidi: base direction from a:pPr@rtl (the parser already flips the default
    // alignment l→r for rtl paragraphs). Engage only when the base is RTL or a
    // line carries strong-RTL characters, so LTR slides keep their exact path.
    const baseRtl = entry.para.rtl === true;
    const paraNeedsBidi = baseRtl || segmentsHaveRtl(line.segments);
    const hasTab = line.segments.some((seg) => seg.isTab);

    if (hasTab) {
      // §21.1.2.1.x: resolve every inline tab in the logical READING frame.
      // UAX#9 L2 later mirrors the resulting cells physically for an RTL base.
      const marLPxE = emuToPx(entry.para.marL, scale);
      const marRPxE = emuToPx(entry.para.marR, scale);
      // The reading pen starts at the leading indent PLUS the draw-side
      // first-line indent (textXOffset): a stop is an ABSOLUTE distance from the
      // leading text-inset edge (§21.1.2.1), so the indent moves where content
      // STARTS, not where a stop sits — omitting it would widen every gap by the
      // indent and slide the cells past their stops. textXOffset only shifts the
      // LTR pen (the RTL draw right-anchors and never applies it), so it joins
      // the LTR frame only.
      const leadingIndentPx = baseRtl ? marRPxE : marLPxE + textXOffset;
      const limitPx = textMaxW + marLPxE + marRPxE;
      const tabFontSeg = line.segments.find((seg) => seg.isTab) as LayoutSegment;
      ctx.font = tabFontSeg.font;
      const spaceW = ctx.measureText(' ').width;
      const items = line.segments.map((seg) => {
        if (seg.isTab) return { isTab: true, width: 0 };
        if (seg.math) return { isTab: false, width: seg.math.width };
        ctx.font = seg.font;
        const ls = seg.letterSpacingPx ?? 0;
        return {
          isTab: false,
          width: seg.text
            ? ctx.measureText(seg.text).width + ls * codePointCount(seg.text)
            : 0,
        };
      });
      const stops = (entry.para.tabStops ?? []).map((stop) => ({
        pos: emuToPx(stop.pos, scale),
        algn: stop.algn,
      }));
      // Default tab grid (§21.1.2.2.7): explicit pPr value, else PowerPoint's
      // universal 1-inch default so an unreached tab snaps to the grid (matching
      // the wrap pass) rather than collapsing to a space (issue #1006).
      const defTabSzPx = emuToPx(entry.para.defTabSz ?? 914400, scale);
      const widths = resolveTabWidths(items, stops, leadingIndentPx, limitPx, spaceW, defTabSzPx);
      for (let i = 0; i < line.segments.length; i++) {
        if (line.segments[i].isTab) line.segments[i].tabWidthPx = widths[i];
      }
    }

    // Measure line for alignment AND baseline ascent in one pass.
    // actualBoundingBoxAscent gives the real font ascent for the rendered glyphs,
    // replacing the 0.8×lineHeight heuristic that over-estimates for CJK and
    // tall fonts, causing text to sit too low within the line box.
    let lineWidth = 0;
    let maxAscent = lineHeight * 0.8; // fallback when no segments
    for (const seg of line.segments) {
      if (seg.isTab) {
        lineWidth += seg.tabWidthPx ?? 0;
        continue;
      }
      if (seg.math) {
        lineWidth += seg.math.width;
        maxAscent = Math.max(maxAscent, seg.math.ascent);
        continue;
      }
      ctx.font = seg.font;
      const m = ctx.measureText(seg.text || 'M');
      const ls = seg.letterSpacingPx ?? 0;
      lineWidth += seg.text ? m.width + ls * codePointCount(seg.text) : 0;
      if (m.actualBoundingBoxAscent > 0) {
        maxAscent = Math.max(maxAscent, m.actualBoundingBoxAscent);
      }
    }
    const baseline = cursorY + maxAscent;

    // Reading-frame marker placement under an RTL base (issue #930, same class as
    // the docx #830 / pptx #913 leading-edge mirroring). PowerPoint seats a list
    // marker's LEADING edge — the RIGHT edge in RTL — at the line's leading
    // (right) edge and lays the text CONTIGUOUSLY to its left, so `text right =
    // leadingEdge − markerAdvance` (verified on the sample PDF: the "1." and "•"
    // markers share one right edge, each text right edge = that edge minus the
    // marker's own width). The prior code mirrored the LTR hanging gutter about
    // the text's right edge — `+ (textX − bulletX)` = `+|indent|` — which pushed
    // the marker |indent| PAST the leading edge into the right margin (the
    // far-frame over-indent this fixes), most visibly for a NARROW bullet whose
    // gap to the text then equalled the whole hanging indent.
    //
    // `leadingEdgeX` is the RTL reading start (= the `algn:'r'` right edge). The
    // marker advance is reserved so the pen below right-aligns the text to
    // `leadingEdgeX − reserve`; 0 for LTR / marker-less lines keeps them
    // byte-identical.
    const leadingEdgeX = textX + textMaxW;
    let rtlMarkerReservePx = 0;
    let rtlBulletBmp: ReturnType<typeof peekCachedBitmapByPath> | null = null;
    if (paraNeedsBidi && baseRtl) {
      if (bulletLabel) {
        ctx.font = bulletFont;
        rtlMarkerReservePx = ctx.measureText(bulletLabel).width;
      } else if (bulletImage && fetchImage) {
        rtlBulletBmp = peekCachedBitmapByPath(bulletImage.imagePath, fetchImage);
        if (rtlBulletBmp) {
          const h = bulletImage.sizePx;
          rtlMarkerReservePx = rtlBulletBmp.height > 0
            ? h * (rtlBulletBmp.width / rtlBulletBmp.height)
            : h;
        }
      }
    }

    // Draw bullet.
    if (bulletLabel) {
      ctx.font = bulletFont;
      ctx.fillStyle = bulletColor;
      if (paraNeedsBidi && baseRtl) {
        const prevDir = ctx.direction;
        ctx.direction = 'rtl';
        // Marker RIGHT edge at the leading edge; its left edge tucks the text
        // that the pen shifts left by `rtlMarkerReservePx` (= this width).
        ctx.fillText(bulletLabel, leadingEdgeX - rtlMarkerReservePx, baseline);
        ctx.direction = prevDir;
      } else {
        ctx.fillText(bulletLabel, bulletX, baseline);
      }
    }

    // Picture bullet (`<a:buBlip>`, ECMA-376 §21.1.2.4.2). The bitmap sits on
    // the text baseline at the same gutter x a char bullet uses. The image was
    // warmed by renderSlide's prefetch pass; if its decode hasn't resolved yet
    // (or fetchImage is absent), draw nothing — the marker simply appears once
    // the bitmap is ready, never blocking the frame.
    if (bulletImage && fetchImage) {
      const bmp = peekCachedBitmapByPath(bulletImage.imagePath, fetchImage);
      if (bmp) {
        // The bullet HEIGHT is the text-derived size (× buSzPct); the WIDTH is
        // derived from the decoded bitmap's intrinsic aspect ratio so a
        // non-square marker isn't squished. §21.1.2.4.2 is silent on the exact
        // dimensions; this mirrors the PowerPoint runtime, which scales the
        // picture to the line text height while preserving its aspect ratio.
        const h = bulletImage.sizePx;
        const w = bmp.height > 0 ? h * (bmp.width / bmp.height) : h;
        const imgY = baseline - h; // bottom-aligned to the baseline
        if (paraNeedsBidi && baseRtl) {
          // Marker RIGHT edge at the leading edge (matching the char-bullet
          // reading-frame placement above); the text pen reserves this width.
          ctx.drawImage(bmp, leadingEdgeX - w, imgY, w, h);
        } else {
          ctx.drawImage(bmp, bulletX, imgY, w, h);
        }
      }
    }

    const effectiveTextX = textX + textXOffset;
    let penX: number;
    if (hasTab) {
      // Tab stops are absolute from the leading text-inset edge; paragraph
      // alignment must not add a second offset (#913/#916).
      penX = baseRtl
        ? textX + textMaxW - rtlMarkerReservePx - lineWidth
        : effectiveTextX;
    } else {
      if (alignment === 'ctr') {
        penX = effectiveTextX + (textMaxW - textXOffset - lineWidth) / 2;
      } else if (alignment === 'r') {
        // Reading-frame (#930): an RTL marker leads at the right edge, so the text
        // right-aligns to `leadingEdge − markerAdvance` (contiguous with the
        // marker). `rtlMarkerReservePx` is 0 for non-list / marker-less lines, so
        // plain RTL paragraphs keep `leadingEdge − lineWidth` (byte-identical).
        penX = textX + textMaxW - rtlMarkerReservePx - lineWidth;
      } else {
        penX = effectiveTextX;
      }
    }

    // Justified alignment (ECMA-376 §20.1.10.59 ST_TextAlignType): just/justLow
    // fill the column by widening inter-word + inter-CJK gaps with the
    // paragraph's last line left natural; dist/thaiDist do the same on every
    // line. Segments are merged by style in layout, so justifyLine splits each
    // line inside its text at the gaps and returns draw pieces carrying `jext`
    // (px to advance after each piece). penX already starts at the left edge
    // (the `else` above). Disabled under bidi (visual≠logical order).
    const justifyMode =
      alignment === 'just' || alignment === 'justLow' ? 'just' as const
      : alignment === 'thaiDist' ? 'thaiDist' as const
      : alignment === 'dist' ? 'dist' as const
      : null;
    // Tab-delimited cells keep their natural widths; the inline gaps provide
    // their stop alignment and must not participate in justification.
    // A `just` line ended by a manual <a:br> is left-aligned like the last line
    // (§20.1.10.59 + §21.1.2.2.1); `dist` ignores this and fills every line
    // (justifyLine only suppresses the last line for `just`).
    const endsLogicalLine = isLastLine || (line.endsWithBreak ?? false);
    const drawSegs = justifyMode && !paraNeedsBidi && !hasTab
      ? justifyLine(line.segments, textMaxW - textXOffset, lineWidth, justifyMode, endsLogicalLine)
      : null;
    const segs: (LayoutSegment & Partial<Justified>)[] = drawSegs ?? line.segments;

    // Visual draw order: under bidi, reorder segments per UAX#9 (rule L2) and
    // draw each with ctx.direction matching its resolved direction. textAlign
    // is already 'left', so penX stays the segment's left edge.
    const visual: LineVisualOrder | null = paraNeedsBidi
      ? computeLineVisualOrder(line.segments, baseRtl)
      : null;
    const segCount = segs.length;
    for (let vi = 0; vi < segCount; vi++) {
      const li = visual ? visual.order[vi] : vi;
      const seg = segs[li];
      const segRtl = visual ? visual.rtl[li] : false;
      if (paraNeedsBidi) ctx.direction = segRtl ? 'rtl' : 'ltr';
      if (seg.isTab) {
        penX += seg.tabWidthPx ?? 0;
        continue;
      }
      // Justification advance after this piece (0 when not justifying). Added to
      // the pen, and folded into the trailing edge of underline / strikethrough
      // and the reported onTextRun width, so decorations and the text layer span
      // the widened gap instead of leaving a hole. The last content piece always
      // has jext 0 (justifyLine never stretches after the final glyph).
      const jext = seg.jext ?? 0;
      const splitBefore = seg.splitBefore;
      const segPerGap = seg.perGap ?? 0;
      const internalStretch =
        splitBefore && splitBefore.length > 0 ? splitBefore.length * segPerGap : 0;

      // ── Equation segment: draw the cached image instead of text ──────────
      if (seg.math) {
        const render = mathRenders.get(seg.math.nodes);
        const w = seg.math.width;
        const h = seg.math.ascent + seg.math.descent;
        if (render && w > 0 && h > 0) {
          const top = baseline - seg.math.ascent;
          const img = tintedMathImage(render, seg.color);
          ctx.drawImage(img, penX, top, w, h);
        }
        penX += w;
        penX += jext;
        continue;
      }
      ctx.font = seg.font;
      ctx.fillStyle = seg.color;
      // baseline shift: OOXML baseline in thousandths of a point; positive = superscript (up)
      const baselineShift = seg.baseline ? -(seg.baseline / 100000) * seg.sizePx : 0;
      const segBaseline = baseline + baselineShift;
      const ls = seg.letterSpacingPx ?? 0;

      // Run-level text highlight (rPr > a:highlight, ECMA-376 §21.1.2.3.4).
      // Box advance = glyph measure + letter spacing + justification stretch
      // (internal CJK pitch + trailing `jext`).
      if (seg.highlight && seg.text) {
        const hlW = ctx.measureText(seg.text).width
          + (ls > 0 ? ls * codePointCount(seg.text) : 0)
          + internalStretch
          + jext;
        paintHighlight(ctx, penX, segBaseline, hlW, seg.sizePx, seg.highlight, seg.color);
      }

      // Run-level text shadow (rPr > effectLst > outerShdw). Set on the
      // context so the fillText below picks it up, then cleared after so
      // the outline / underline / strikethrough don't get shadowed too.
      // ECMA-376 §20.1.8.45 dir is degrees clockwise from east; the same
      // formula used by the shape-level shadow renderer above.
      const segShadow = seg.shadow;
      if (segShadow) {
        const dirRad = (segShadow.dir * Math.PI) / 180;
        const dist = emuToPx(segShadow.dist, scale);
        ctx.save();
        ctx.shadowColor = hexToRgba(segShadow.color, segShadow.alpha);
        ctx.shadowBlur = emuToPx(segShadow.blur, scale);
        ctx.shadowOffsetX = Math.cos(dirRad) * dist;
        ctx.shadowOffsetY = Math.sin(dirRad) * dist;
      }

      // Draw `text` at `atX` honouring this segment's letter-spacing / RTL
      // semantics. Lifted into a closure so the fill and stroke (outline) paths
      // share it, and so the split-CJK branch below can call it per piece.
      const drawWithFont = (text: string, atX: number, op: 'fill' | 'stroke'): void => {
        const paint = op === 'fill' ? ctx.fillText.bind(ctx) : ctx.strokeText.bind(ctx);
        if (ls > 0 && text.length > 1) {
          // rPr @spc (§21.1.2.3.x): distribute the per-glyph advance via canvas
          // letterSpacing and draw the whole CONTEXTUALLY-shaped string in ONE
          // paint. This keeps the drawn advance == the layout's contextual
          // measure + n·ls (約物半角 contextual half-width collapse honoured, so a
          // piece's glyphs stay aligned with its contextual `dx` origin and the
          // next piece/run never overlaps — the pptx analog of docx PR #626), AND
          // preserves Arabic cursive joining for RTL. (The old LTR branch summed
          // ISOLATED measure(ch)+ls, which overran the contextual box at 約物
          // punctuation.) Chromium's measureText adds letterSpacing after every
          // glyph incl. the trailing one (= natural + n·ls), matching
          // codePointCount(seg.text)·ls used by the layout's segW.
          const lctx = ctx as CanvasRenderingContext2D & { letterSpacing: string };
          const prev = lctx.letterSpacing;
          try { lctx.letterSpacing = `${ls}px`; } catch { /* older engines */ }
          paint(text, atX, segBaseline);
          try { lctx.letterSpacing = prev; } catch { /* ignore */ }
        } else {
          paint(text, atX, segBaseline);
        }
      };

      // `splitBefore` is set when justifyLine annotated this segment with
      // internal CJK gaps. Anchor each piece to the whole-string prefix advance
      // (via core's `justifiedPiecePositions`) to absorb 約物半角 drift — see
      // packages/core/src/text/justify-positions.ts for the invariant proof.
      const measureCb = (s: string): number => ctx.measureText(s).width;
      const pieces =
        splitBefore && splitBefore.length > 0
          ? justifiedPiecePositions([...seg.text], splitBefore, segPerGap, measureCb, ls)
          : null;

      // A run is FULLY distributed when justifyLine opened a gap at EVERY internal
      // inter-glyph boundary (splitBefore.length === cps.length - 1) — e.g. a
      // pure-CJK `algn="just"/"dist"` line. Drawing one single-glyph piece per code
      // point (the `pieces` loop) routes through drawWithFont's `else → paint(ch)`:
      // each piece has text.length === 1, so the `ls > 0 && text.length > 1`
      // letterSpacing branch never fires, and every glyph is painted ISOLATED even
      // when ls > 0. That loses the browser's JIS X 4051 約物連続 packing — a
      // closing-class punct immediately followed by an opening bracket ("：［",
      // "、［", "）（") packs ~half-em in measureText (a bracket next to a plain
      // kanji/kana does NOT pack). The isolated full-width bracket then overruns its
      // successor (the justify analog of docx PR #630). The fix: a gap at EVERY
      // boundary ⇒ uniform per-glyph pitch (ls + segPerGap), so draw the whole
      // CONTEXTUALLY-shaped run in ONE paint with ctx.letterSpacing = ls+segPerGap:
      // glyph i lands at measure(prefix_i)+i·(ls+segPerGap) — the exact justified
      // position — so the packing is honoured and nothing overlaps; the final glyph
      // reaches the segment box edge (segW). measureCb already ran (pieces built
      // above) at the natural advance, so no measureText sees this letterSpacing.
      const cps = [...seg.text];
      const fullyDistributed =
        !!splitBefore && splitBefore.length === cps.length - 1 && cps.length > 1;
      const drawRun = (op: 'fill' | 'stroke'): void => {
        if (eaVertUpright) {
          // ECMA-376 §20.1.10.83 eaVert: paint each glyph with its UAX#50 vertical
          // orientation (CJK/kana upright, Latin sideways, brackets/comma
          // substituted, ー rotated). The segment's cell positions are already set
          // by the horizontal layout (penX / segW), so only the glyph orientation
          // changes. A FULLY-DISTRIBUTED line (algn just/dist opening a gap at every
          // inter-glyph boundary — e.g. 均等割り付け of a pure-CJK column) carries its
          // uniform justification pitch in `segPerGap`; fold it into the per-glyph
          // advance so the column SPREADS to fill the length instead of bunching at
          // the top (mirrors the horizontal fully-distributed fast-path below).
          // Partial `just` distribution (mixed Latin/CJK pieces) is a rare eaVert
          // edge and is not redistributed here.
          const eaPitch = fullyDistributed ? ls + segPerGap : ls;
          drawEaVertRun(ctx, seg.text, penX, segBaseline, seg.sizePx, eaPitch, op);
          return;
        }
        if (fullyDistributed) {
          const lctx = ctx as CanvasRenderingContext2D & { letterSpacing: string };
          const prev = lctx.letterSpacing;
          try { lctx.letterSpacing = `${ls + segPerGap}px`; } catch { /* older engines */ }
          (op === 'fill' ? ctx.fillText.bind(ctx) : ctx.strokeText.bind(ctx))(
            seg.text,
            penX,
            segBaseline,
          );
          try { lctx.letterSpacing = prev; } catch { /* ignore */ }
        } else if (pieces) {
          for (const { text: pieceText, dx } of pieces) drawWithFont(pieceText, penX + dx, op);
        } else {
          drawWithFont(seg.text, penX, op);
        }
      };

      drawRun('fill');

      if (segShadow) ctx.restore();

      // Run-level text outline (rPr > a:ln). Strokes each glyph in addition
      // to the fill so the text reads as a thin lined character. ECMA-376
      // §20.1.2.2.24: `w` is EMU; convert to px via the same scale used for
      // fonts. Min 0.5 px so a 1-pt outline at small zoom levels stays
      // visible. Skip when width <= 0 (parser may emit width=0 for
      // explicitly empty <a:ln/>).
      const segOutline = seg.outline;
      if (segOutline && segOutline.width > 0) {
        ctx.save();
        ctx.lineWidth = Math.max(0.5, emuToPx(segOutline.width, scale));
        ctx.strokeStyle = segOutline.color ? `#${segOutline.color}` : seg.color;
        ctx.lineJoin = 'round';
        drawRun('stroke');
        ctx.restore();
      }

      ctx.font = seg.font;
      const baseW = ctx.measureText(seg.text).width;
      const segW = baseW + (ls > 0 ? ls * codePointCount(seg.text) : 0) + internalStretch;

      if (onTextRun && seg.text) {
        onTextRun({
          text: seg.text,
          inShapeX: penX - bx,
          inShapeY: cursorY - by,
          w: segW + jext,
          h: lineHeight,
          fontSize: seg.sizePx,
          font: seg.font,
          shapeX: bx,
          shapeY: by,
          shapeW: bw,
          shapeH: bh,
          rotation: shapeRotation,
          hyperlink: seg.hyperlink,
        });
      }

      if (seg.underline) {
        drawUnderline(ctx, penX, segBaseline, segW + jext, seg.sizePx, seg.underlineColor ?? seg.color, seg.underlineStyle, rc.dpr);
      }

      if (seg.strikethrough) {
        const lineW = Math.max(1, seg.sizePx * 0.05);
        ctx.strokeStyle = seg.color;
        ctx.lineWidth = lineW;
        ctx.setLineDash([]);
        // Crispness nudge (see crispOffset): the strike is a horizontal stroke;
        // an odd device-pixel width is centered on one device row by snapping its
        // y onto the nearest crisp device position (otherwise it straddles two
        // rows → blurry). Snap each line from its own y.
        const yMid = segBaseline - seg.sizePx * 0.32;
        if (seg.strikeDouble) {
          // Two parallel lines straddling the standard strike position,
          // separated by ~ 1.5× the line weight (visually distinct yet
          // staying inside the glyph's central band).
          const offset = lineW * 0.9;
          const yUp = yMid - offset;
          const yDn = yMid + offset;
          ctx.beginPath();
          ctx.moveTo(penX, yUp + crispOffset(yUp, lineW, rc.dpr));
          ctx.lineTo(penX + segW + jext, yUp + crispOffset(yUp, lineW, rc.dpr));
          ctx.moveTo(penX, yDn + crispOffset(yDn, lineW, rc.dpr));
          ctx.lineTo(penX + segW + jext, yDn + crispOffset(yDn, lineW, rc.dpr));
          ctx.stroke();
        } else {
          const sy = yMid + crispOffset(yMid, lineW, rc.dpr);
          ctx.beginPath();
          ctx.moveTo(penX, sy);
          ctx.lineTo(penX + segW + jext, sy);
          ctx.stroke();
        }
      }

      penX += segW;
      penX += jext;
    }
    if (paraNeedsBidi) ctx.direction = 'ltr';

    cursorY += linePx;
  }

  ctx.restore();
}

// The lazy image-byte source closure (one stable identity per Presentation, so
// it namespaces the shared decoded-bitmap cache per deck).
type FetchImage = (path: string, mime: string) => Promise<Blob>;

// The decoded raster/metafile bitmap cache now lives in core
// (`getCachedBitmapByPath` / `peekCachedBitmapByPath` / `dropBitmapCacheByPath`),
// shared verbatim with docx and xlsx. Re-exported under the historical pptx
// names so the presentation teardown and the bullet-draw tests keep their import
// surface; the synchronous picture-bullet draw reads a warmed bitmap through
// `peekCachedBitmap`. See core/src/image/bitmap-image-by-path.ts.
export {
  getCachedBitmapByPath as getCachedBitmap,
  peekCachedBitmapByPath as peekCachedBitmap,
  dropBitmapCacheByPath as dropImageBitmapCache,
  // Second-layer duotone (§20.1.8.23) recolour cache; dropped alongside the base
  // bitmap cache on deck teardown.
  dropDuotoneBitmapCache,
} from '@silurus/ooxml-core';


/** Local view of the parsed `<a:sp3d>` (1:1 with the Sp3d TS type). */
interface Sp3dLike {
  extrusionH?: number;
  extrusionClr?: string;
  bevelT?: { w: number; h: number; prst: string };
  bevelB?: { w: number; h: number; prst: string };
}
/** Local view of `<a:lightRig>` (1:1 with the LightRig TS type). */
interface LightRigLike {
  rig: string;
  dir: string;
  /** Optional `<a:rot lat lon rev>` (§20.1.5.11) — rev rotates the key azimuth. */
  rot?: { lat: number; lon: number; rev: number };
}

/**
 * Build the bevel-shading inputs for an sp3d, in DEVICE px. ECMA-376 §20.1.5.12
 * `bevelT`/`bevelB`. Returns an empty list when there is no bevel to shade.
 *
 * The bevel band/height are EMU (§20.1.5.3); convert to device px via the EMU→
 * CSS-px `scale` and the canvas `devScale`, because the shading reads/writes the
 * device-resolution offscreen. The light vector comes from the scene's lightRig
 * (§20.1.5.9); when absent we default to the OOXML default rig "threePt"/"t"
 * (PowerPoint's default 3-D scene light) so a bevel with no explicit rig still
 * picks up a top key light rather than rendering flat.
 */
export function buildBevelInputs(
  sp3d: Sp3dLike | undefined,
  lightRig: LightRigLike | undefined,
  prstMaterial: string | undefined,
  scale: number,
  devScale: number,
): BevelInput[] {
  if (!sp3d) return [];
  const rig = lightRig?.rig ?? 'threePt';
  const dir = lightRig?.dir ?? 't';
  // §20.1.5.9/§20.1.5.11: the lightRig's `<a:rot>` (rev = in-plane revolution about
  // the view axis) rotates the key-light azimuth. lat/lon are a documented SPEC GAP
  // in lightDirFromRig (no calibration sample exercises them).
  const light = lightDirFromRig(rig, dir, lightRig?.rot);
  const material = materialClass(prstMaterial);
  const emuToDev = scale * devScale;
  const out: BevelInput[] = [];
  if (sp3d.bevelT && sp3d.bevelT.w > 0 && sp3d.bevelT.h > 0) {
    out.push({
      widthPx: sp3d.bevelT.w * emuToDev,
      heightPx: sp3d.bevelT.h * emuToDev,
      prst: sp3d.bevelT.prst || 'circle',
      material,
      light,
    });
  }
  if (sp3d.bevelB && sp3d.bevelB.w > 0 && sp3d.bevelB.h > 0) {
    out.push({
      widthPx: sp3d.bevelB.w * emuToDev,
      heightPx: sp3d.bevelB.h * emuToDev,
      prst: sp3d.bevelB.prst || 'circle',
      material,
      light,
      bottom: true,
    });
  }
  return out;
}

/**
 * Build the extrusion side-wall input (device px) for an sp3d, or null when
 * there is no extrusion or the camera is face-on (the wall would be invisible).
 * ECMA-376 §20.1.5.12 `extrusionH` / `extrusionClr`. The side-wall colour is the
 * explicit `extrusionClr` when present, else a darkened mid-grey stand-in (the
 * spec leaves the default to the renderer; PowerPoint uses a shaded copy of the
 * face — a neutral dark wall is the conservative Phase B choice).
 */
function buildExtrusion(
  sp3d: Sp3dLike | undefined,
  camera: CameraInput,
  w: number,
  h: number,
  scale: number,
  devScale: number,
): ExtrusionInput | null {
  if (!sp3d || !sp3d.extrusionH || sp3d.extrusionH <= 0) return null;
  const depthDev = sp3d.extrusionH * scale * devScale;
  // Depth offset is computed in the device-px box (w·devScale × h·devScale) so
  // it matches the device-resolution offscreen the side wall is baked into.
  const off = computeDepthOffset(camera, w * devScale, h * devScale, depthDev);
  if (Math.hypot(off.x, off.y) < 0.75) return null;
  let rgb: [number, number, number] = [64, 64, 64];
  if (sp3d.extrusionClr) {
    const hex = sp3d.extrusionClr.replace('#', '');
    if (hex.length >= 6) {
      rgb = [
        parseInt(hex.slice(0, 2), 16),
        parseInt(hex.slice(2, 4), 16),
        parseInt(hex.slice(4, 6), 16),
      ];
    }
  }
  return { offsetX: off.x, offsetY: off.y, rgb };
}

/**
 * Project a shape/picture body through a DrawingML 3D camera (scene3d, Phase A),
 * with optional sp3d bevel shading baked in BEFORE the projection (Phase B).
 *
 * ECMA-376 §20.1.5.5: the camera acts in the shape's *local* space, ahead of
 * the shape's 2D placement. We therefore render the body — normally drawn at
 * (x,y,w,h) in `target`'s coordinate space — into a separate device-resolution
 * offscreen canvas sized to the w×h box, then warp that bitmap onto `target`
 * with `drawProjected` (a piecewise-affine homography). Because `target` already
 * carries the element's 2D rotation/flip transform when this runs, the
 * projection composes *inside* that transform, exactly matching "scene3d first,
 * then the 2D xfrm" ordering.
 *
 * Phase B: when `bevels` is non-empty, the bevel lip shading is baked into the
 * offscreen bitmap (`applyBevelShading`) right after `paintBody` and before the
 * warp, so the lit/shadowed rim rides the same camera homography as the body
 * (matching PowerPoint, which shades the 3-D solid then projects it).
 *
 * `paintBody(octx, ox, oy, ow, oh)` paints the body (clip + fill/stroke/image)
 * into the offscreen context at the given local rect. Returns true if it
 * projected, false when no offscreen canvas is available (headless) — callers
 * then fall back to painting directly.
 */
interface Project3dOpts {
  /** Bevel lips to bake into the body before the warp (§20.1.5.12 bevelT/B). */
  bevels?: BevelInput[];
  /** Extrusion side-wall to bake in before the bevel (§20.1.5.12 extrusionH). */
  extrusion?: ExtrusionInput;
  /**
   * Paint edges that sit OUTSIDE the beveled front face (the sp3d contour, an
   * outside-aligned rim of the extruded side edge) into the offscreen AFTER the
   * bevel shading, in the same local rect. The a:ln border belongs to the front
   * face itself — PowerPoint's bevel lights the framed picture as one surface —
   * so the border is painted inside `paintBody`, BEFORE the bevel, and only the
   * contour goes here (it must not contaminate the lip's distance transform).
   */
  paintEdges?: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void;
  /**
   * Margin (CSS px) added on every side of the offscreen around the shape's
   * w×h box. The box alone CLIPS everything a shape legitimately paints past
   * it: the outer half of the centre-aligned a:ln border, the outside-aligned
   * contour, and the extrusion sweep. For a silhouette that touches its box
   * (an inscribed ellipse touches it at all four apices) that clip cuts the
   * rim along the straight box edge — the "sliced ellipse" slide-6 bug. The
   * margin also keeps the silhouette away from the offscreen border so the
   * bevel's distance transform (which treats out-of-bounds as background)
   * measures the true silhouette distance instead of the canvas edge.
   */
  edgePadCss?: number;
}

function projectScene3dPaint(
  target: CanvasRenderingContext2D,
  camera: CameraInput,
  x: number,
  y: number,
  w: number,
  h: number,
  paintBody: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  opts: Project3dOpts = {},
): boolean {
  if (w <= 0 || h <= 0) return false;
  // Device scale folded into target's current transform, so the offscreen is
  // rasterised at the same pixel density as the live canvas (no blur on HiDPI).
  const tf = target.getTransform();
  const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;

  // Edge margin (see Project3dOpts.edgePadCss). Quantised to whole device px so
  // the body lands on the same pixel grid as the unpadded layout, then mapped
  // onto the destination by extrapolating the camera quad through its own
  // homography (computeScene3dQuad re-fits to the box it is given, so it must
  // NOT be called with the padded size). Degenerate extrapolation → pad 0.
  let padDev = Math.max(0, Math.ceil((opts.edgePadCss ?? 0) * devScale));
  const quad = computeScene3dQuad(camera, w, h);
  let quadCorners = quad.corners;
  if (padDev > 0) {
    const padCss = padDev / devScale;
    const expanded = expandProjectedQuad(quad.corners, padCss / w, padCss / h);
    if (expanded) quadCorners = expanded;
    else padDev = 0;
  }
  const padCss = padDev / devScale;

  const ow = Math.max(1, Math.ceil(w * devScale) + 2 * padDev);
  const oh = Math.max(1, Math.ceil(h * devScale) + 2 * padDev);
  const aux = createAuxCanvas(ow, oh);
  if (!aux) return false;
  const octx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!octx) return false;

  // Paint the body into the offscreen with the origin at (pad,pad); device px.
  octx.save();
  octx.scale(devScale, devScale);
  octx.translate(padCss, padCss);
  paintBody(octx, 0, 0, w, h);
  octx.restore();

  // The body silhouette occupies the offscreen's inner box (device px):
  // [padDev, padDev + w·devScale] on each axis (the a:ln border, painted inside
  // paintBody, extends ±half-line-width past it; the pad reserves that room).
  // Restricting the bevel/extrusion distance-transform + sweep to this box grown
  // by the effect's reach (perf: A3) skips the transparent pad border. Equivalent
  // to the whole offscreen because a band pixel is within `bandPx` of the
  // silhouette (lip inside it) and a wall pixel within `|offset|` of it. When the
  // offscreen is already shape-tight (the usual case) the region ≈ the whole
  // canvas and this is a harmless no-op. See BevelRegion.
  const bodyDevW = Math.ceil(w * devScale);
  const bodyDevH = Math.ceil(h * devScale);
  const regionFor = (reachPx: number): BevelRegion => ({
    x: padDev - reachPx,
    y: padDev - reachPx,
    w: bodyDevW + 2 * reachPx,
    h: bodyDevH + 2 * reachPx,
  });

  // Extrusion side wall first (it sits behind/around the front face), then the
  // bevel lip on the front-face silhouette (§20.1.5.12). Both are baked into the
  // offscreen so they ride the camera warp.
  if (opts.extrusion) {
    const reach = Math.ceil(Math.hypot(opts.extrusion.offsetX, opts.extrusion.offsetY)) + 2;
    applyExtrusion(octx, opts.extrusion, regionFor(reach));
  }
  // Bake bevel lip shading into the body bitmap (device-px band) before the
  // warp so the lit/shadowed rim rides the camera homography (§20.1.5.12).
  if (opts.bevels && opts.bevels.length > 0) {
    for (const bevel of opts.bevels) {
      applyBevelShading(octx, bevel, regionFor(Math.ceil(bevel.widthPx) + 2));
    }
  }

  // Post-bevel edges (the contour rim) on top of the beveled body, still
  // pre-projection. The a:ln border is painted inside paintBody (see
  // Project3dOpts.paintEdges).
  if (opts.paintEdges) {
    octx.save();
    octx.scale(devScale, devScale);
    octx.translate(padCss, padCss);
    opts.paintEdges(octx, 0, 0, w, h);
    octx.restore();
  }

  // Offset the (possibly pad-expanded) quad to the element's (x,y).
  const corners = quadCorners.map((c) => ({ x: x + c.x, y: y + c.y })) as [
    Vec2,
    Vec2,
    Vec2,
    Vec2,
  ];
  // drawProjected runs in target's CSS-px space (the device scale lives in the
  // target transform), warping the ow×oh offscreen onto the CSS-space quad.
  drawProjected(aux as unknown as CanvasImageSource, target, ow, oh, corners);
  return true;
}

/**
 * Bake bevel shading + edges into a body that is NOT camera-projected (identity
 * camera, e.g. orthographicFront). Renders the body to a device-px offscreen,
 * shades the bevel, paints the post-bevel edges (contour), then blits the
 * offscreen back at (x,y). Falls back to false (caller paints flat) when no
 * offscreen is available.
 *
 * `edgePadCss` grows the offscreen by a margin on every side (same rationale
 * as Project3dOpts.edgePadCss): without it the box-sized offscreen clips the
 * outer half of the centre-aligned border wherever the silhouette touches its
 * bounding box, and the bevel's distance transform reads the offscreen border
 * as the silhouette edge there — both showed up as the straight-cut rim on the
 * slide-6 ellipse apices.
 */
function paintBeveledFlat(
  target: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  bevels: BevelInput[],
  paintBody: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  paintEdges?: (octx: CanvasRenderingContext2D, ox: number, oy: number, ow: number, oh: number) => void,
  edgePadCss = 0,
): boolean {
  if (w <= 0 || h <= 0 || bevels.length === 0) return false;
  const tf = target.getTransform();
  const det = Math.abs(tf.a * tf.d - tf.b * tf.c);
  const devScale = det > 0 ? Math.sqrt(det) : 1;
  // Whole-device-px margin so the body stays on the unpadded pixel grid.
  const padDev = Math.max(0, Math.ceil(edgePadCss * devScale));
  const padCss = padDev / devScale;
  const ow = Math.max(1, Math.ceil(w * devScale) + 2 * padDev);
  const oh = Math.max(1, Math.ceil(h * devScale) + 2 * padDev);
  const aux = createAuxCanvas(ow, oh);
  if (!aux) return false;
  const octx = aux.getContext('2d') as CanvasRenderingContext2D | null;
  if (!octx) return false;

  octx.save();
  octx.scale(devScale, devScale);
  octx.translate(padCss, padCss);
  paintBody(octx, 0, 0, w, h);
  octx.restore();
  // Restrict the bevel distance-transform to the body's inner box grown by the
  // band width (perf: A3 — skips the transparent pad border). Equivalent because
  // the lip sits within `bandPx` of the silhouette, which fills this box. No-op
  // when the offscreen is already shape-tight. See BevelRegion / projectScene3dPaint.
  const bodyDevW = Math.ceil(w * devScale);
  const bodyDevH = Math.ceil(h * devScale);
  for (const bevel of bevels) {
    const reach = Math.ceil(bevel.widthPx) + 2;
    applyBevelShading(octx, bevel, {
      x: padDev - reach,
      y: padDev - reach,
      w: bodyDevW + 2 * reach,
      h: bodyDevH + 2 * reach,
    });
  }
  if (paintEdges) {
    octx.save();
    octx.scale(devScale, devScale);
    octx.translate(padCss, padCss);
    paintEdges(octx, 0, 0, w, h);
    octx.restore();
  }
  // Blit the device-px offscreen back into target's CSS-px space, 1:1 in
  // device pixels, with the margin hanging past (x,y) on every side.
  target.drawImage(
    aux as unknown as CanvasImageSource,
    x - padCss,
    y - padCss,
    ow / devScale,
    oh / devScale,
  );
  return true;
}

/** Poster bitmaps decoded once per media element; renderSlide's prefetch pass
 *  warms this so the sequential draw loop never waits on the network. Keyed by
 *  element identity (not posterPath), so the bitmap releases when the slide
 *  model is GC'd; the same poster on two elements decodes twice, which is
 *  bounded and fine for the per-slide warm-up this serves. */
const posterBitmapCache = new WeakMap<MediaElement, Promise<ImageBitmap>>();

// Exported for the RB1 poster decode-bomb neutralization test (asserts an
// over-budget poster is rejected before `createImageBitmap`). Not part of the
// public package surface.
export function getPosterBitmap(
  el: MediaElement,
  fetchMedia: (path: string) => Promise<Blob>,
): Promise<ImageBitmap> {
  const hit = posterBitmapCache.get(el);
  if (hit) return hit;
  const p = (async () => {
    const blob = await fetchMedia(el.posterPath);
    const typed = el.posterMimeType
      ? new Blob([blob], { type: el.posterMimeType })
      : blob;
    // Decode-bomb guard: `el.posterPath`/`posterMimeType` come from the
    // attacker-controllable `<a:blip>` (shape.rs), and `createImageBitmap` sizes
    // its decoded RGBA surface from the image HEADER, not the compressed length —
    // a tiny PNG/JPEG declaring e.g. 60000×60000 forces a multi-GB allocation.
    // Sniff the pixel dimensions from a 64 KiB header prefix and reject an
    // over-budget raster BEFORE it reaches createImageBitmap, exactly as
    // `decodeRasterOrMetafile` (RB1) does for picture blips. A rejection here
    // makes both callers fall through to their plain media fill (graceful
    // degradation). The 64 KiB prefix covers a JPEG SOF past EXIF/ICC; an
    // unrecognized header is not blocked (fail-open).
    const head = new Uint8Array(await typed.slice(0, 64 * 1024).arrayBuffer());
    if (rasterHeaderExceedsBudget(head)) {
      throw new Error('poster raster exceeds the pixel budget');
    }
    return createImageBitmap(typed);
  })();
  posterBitmapCache.set(el, p);
  return p;
}

async function renderPicture(
  ctx: CanvasRenderingContext2D,
  el: PictureElement,
  scale: number,
  fetchImage?: (path: string, mime: string) => Promise<Blob>,
) {
  // No byte source → nothing to draw (the lazy pipeline always supplies one in
  // both render modes; this guards the rare misconfiguration).
  if (!fetchImage) return;
  try {
    // Prefer the vector original (Microsoft svgBlip extension); fall back to the
    // raster on any SVG decode failure. `bitmap` widens to the union of the two
    // drawable sources — both expose numeric .width/.height (used for the
    // srcRect crop below) and both are valid `ctx.drawImage` sources, so every
    // downstream path stays unchanged.
    // `imagePath` is normally a raster (PNG/JPEG), but for a pure-SVG picture
    // with no raster blip it is the SVG part itself — and `createImageBitmap`
    // (getCachedBitmapByPath) cannot rasterize SVG in every browser, so such a
    // picture must also go through the <img>-based SVG decoder (keyed by path).
    const dataIsSvg = el.mimeType === 'image/svg+xml';
    // The picture's intended draw size in points sizes any metafile raster
    // (el.width/height are EMU; /PT_TO_EMU = pt). Unused by the raster/SVG paths.
    // A cropped metafile rasterizes at its FULL picture frame (scaled up by
    // 1/(1−crop)) so the fractional crop below lands correctly; raster blips and
    // uncropped metafiles pass through unchanged. NB: getCachedBitmapByPath is keyed by
    // imagePath ("first size wins"), so if one path is referenced both cropped
    // and uncropped on a slide only the first decode's raster size is kept — that
    // affects raster SHARPNESS only; the crop fraction itself is applied per
    // element from el.srcRect at draw time, so geometry stays correct either way.
    const { widthPt, heightPt } = metafileRasterSize(
      el.mimeType,
      el.srcRect,
      el.width / PT_TO_EMU,
      el.height / PT_TO_EMU,
    );
    // `null` is reachable when the raster path resolves to an unsupported
    // metafile (a true EMF, or a WMF with no geometry); guarded below.
    let bitmap: ImageBitmap | HTMLImageElement | null;
    if (preferVectorBlip(el)) {
      // No crop: prefer the vector original (shared gate; see preferVectorBlip).
      // With an a:srcRect crop we skip this branch — the crop math below
      // multiplies fractional srcRect edges by the source's pixel dims, and an
      // SVG HTMLImageElement that declares only a viewBox (no intrinsic
      // width/height) reports the 300×150 default rather than its logical size,
      // so a 9-arg drawImage with a source rect samples the wrong basis. When a
      // real raster exists it has exact pixel dims that make the ECMA-376
      // §20.1.8.55 fractional crop well-defined (handled in the final branch);
      // when only the SVG exists the `dataIsSvg` branch below still draws it
      // (uncropped is the overwhelmingly common case for icons).
      try {
        bitmap = await getCachedSvgImageByPath(el.svgImagePath, fetchImage);
      } catch {
        // §20.1.8.23 duotone recolour applies only to the raster fallback — an
        // SVG vector original has no readable pixel grid (matches xlsx).
        bitmap = dataIsSvg
          ? await getCachedSvgImageByPath(el.imagePath, fetchImage)
          : await getCachedDuotoneBitmapByPath(el.imagePath, el.mimeType, el.duotone, fetchImage, { widthPt, heightPt });
      }
    } else if (dataIsSvg) {
      // SVG-only picture (here either because it has a crop, or — defensively —
      // because no svgImagePath was surfaced): decode through the SVG path since
      // createImageBitmap can't. A duotone on a vector picture is a rare edge
      // case left un-recoloured (no readable pixel grid), matching xlsx.
      bitmap = await getCachedSvgImageByPath(el.imagePath, fetchImage);
    } else {
      // §20.1.8.23 duotone recolour on the raster blip (once, at decode time,
      // cached under a colour-suffixed key). No duotone ⇒ this is exactly the
      // former `getCachedBitmapByPath` decode.
      bitmap = await getCachedDuotoneBitmapByPath(el.imagePath, el.mimeType, el.duotone, fetchImage, { widthPt, heightPt });
    }
    // Skip a picture whose blip is an unsupported metafile (null bitmap), the
    // same way an SVG-decode failure that also fails its raster fallback would
    // throw out of this try — here we simply return without painting.
    if (!bitmap) return;
    ctx.save();
    if (el.alpha != null) ctx.globalAlpha *= el.alpha;
    const x = emuToPx(el.x, scale);
    const y = emuToPx(el.y, scale);
    const w = emuToPx(el.width, scale);
    const h = emuToPx(el.height, scale);
    if (el.rotation !== 0 || el.flipH || el.flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      ctx.rotate((el.rotation * Math.PI) / 180);
      if (el.flipH) ctx.scale(-1, 1);
      if (el.flipV) ctx.scale(1, -1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }

    // ── srcRect sub-rectangle (ECMA-376 §20.1.8.55 a:srcRect). Resolve once via
    // the shared core helper so both the live paint and the effect aux paints
    // share identical crop coordinates. Applies to raster blips AND metafiles
    // alike: a cropped metafile was rasterized at its full picture frame above
    // (`metafileRasterSize`), so its bitmap maps to the same fractional source.
    const crop = cropSourceRect(bitmap, el.srcRect);

    // Trace the picture's clip silhouette (roundRect / custGeom / plain rect).
    // Shared by the clip and the border / contour strokes so the outline always
    // hugs the exact silhouette the bitmap is trimmed to. ECMA-376 §20.1.9.8:
    // a `<p:pic>` may carry `<a:custGeom>` (e.g. a laptop frame) or a roundRect
    // preset clip.
    //
    // `...Subpath` appends the silhouette as a fresh subpath of the CURRENT path
    // (no `beginPath`). Used when combining the silhouette with an enclosing
    // rect for an even-odd "outside" clip region.
    const tracePictureSilhouetteSubpath = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      if (el.custGeom && el.custGeom.length > 0) {
        // custGeom takes priority over prstGeom (ECMA-376 §20.1.9.8).
        buildCustomPath(target, el.custGeom, cx, cy, cw, ch);
      } else if (el.prstGeom) {
        // §20.1.9.18 — the picture's preset geometry is its clip silhouette.
        // Driven by the shared preset-geometry engine (roundRect, ellipse, and
        // the other 184 presets), with the avLst adjust handles; the engine
        // substitutes each preset's declared default for any omitted guide.
        const ok = buildPresetGeometryPath(
          target,
          el.prstGeom,
          cx,
          cy,
          cw,
          ch,
          el.prstAdjust ?? [],
        );
        // Unknown preset name → fall back to a plain rectangle so the bitmap
        // still draws (matches the pre-generalisation rect fallback).
        if (!ok) target.rect(cx, cy, cw, ch);
      } else {
        target.rect(cx, cy, cw, ch);
      }
    };

    const tracePictureSilhouette = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      target.beginPath();
      tracePictureSilhouetteSubpath(target, cx, cy, cw, ch);
    };

    // Apply the picture clip (roundRect / custGeom) at an arbitrary local rect.
    // The rect is parameterised so the same clip applies to the live draw
    // (x,y,w,h) and to the scene3d offscreen (0,0,w,h). A plain rectangle needs
    // no clip (the bitmap already fills the rect), so we only clip when there
    // is an actual non-rectangular / rounded silhouette.
    const applyClipAt = (
      target: CanvasRenderingContext2D,
      cx: number,
      cy: number,
      cw: number,
      ch: number,
    ): void => {
      if (el.prstGeom || (el.custGeom && el.custGeom.length > 0)) {
        tracePictureSilhouette(target, cx, cy, cw, ch);
        target.clip();
      }
    };

    // Stroke the picture's silhouette with the `<a:ln>` border and/or the
    // `<a:sp3d>` contour edge. Drawn *after* the bitmap inside `paintImageAt`,
    // so when scene3d is active the strokes are warped through the same camera
    // homography and effectLst (reflection / soft edge) re-paints them too —
    // matching PowerPoint, which applies the 3D transform and effects to the
    // framed picture as a whole.
    //
    // The two edges are split because they sit on OPPOSITE sides of the bevel
    // shading: the a:ln border is part of the FRONT FACE (PowerPoint's bevel
    // lip lights the framed picture as one surface — sample-11.pdf p6 shows
    // the lit shelf and dark rim crease ON the beige border), so it is painted
    // before applyBevelShading; the contour approximates the extruded side
    // edge OUTSIDE the front face and must stay out of the lip's distance
    // transform, so it is painted after.
    const strokeLnBorder = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      // a:ln picture border (ECMA-376 §20.1.2.2.24). Centre-aligned stroke,
      // the Canvas default — PowerPoint draws the picture frame straddling
      // the silhouette edge.
      if (el.stroke) {
        target.save();
        applyStroke(target, el.stroke, scale);
        tracePictureSilhouette(target, ox, oy, ow, oh);
        target.stroke();
        target.restore();
      }
    };
    const strokeContourEdge = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      // sp3d contour edge (ECMA-376 §20.1.5.12 `contourW` / `<a:contourClr>`).
      //    The spec's contour is the extruded 3D edge surface lit by the scene's
      //    light rig. We draw a FLAT approximation: a uniform-width outline in
      //    the contour colour, with no per-edge light-rig response. (The bevel
      //    lip itself IS shaded — applyBevelShading runs on the body before this;
      //    the contour is the thin rim OUTSIDE that bevel.) Position assumption:
      //    the contour grows OUTWARD from the front face in 3D, so we draw it as
      //    an OUTSIDE-aligned stroke (the framed edge sits just beyond the
      //    picture, not over the image). Canvas has no outside-stroke mode, so we
      //    (a) clip to the region OUTSIDE the silhouette — traced silhouette + an
      //    enclosing rect with the even-odd rule — then (b) stroke the silhouette
      //    at 2× width centred on the edge; only the outer half survives the clip.
      const sp3d = el.sp3d;
      if (sp3d && (sp3d.contourW ?? 0) > 0 && sp3d.contourClr) {
        const wPx = Math.max(0.5, (sp3d.contourW as number) * scale);
        target.save();
        // Clip to everything OUTSIDE the silhouette: trace the silhouette plus a
        // generously enlarged enclosing rect, filled even-odd, so the silhouette
        // interior is excluded from the clip region.
        target.beginPath();
        const pad = wPx * 2 + Math.max(ow, oh);
        target.rect(ox - pad, oy - pad, ow + 2 * pad, oh + 2 * pad);
        tracePictureSilhouetteSubpath(target, ox, oy, ow, oh);
        target.clip('evenodd');
        target.beginPath();
        tracePictureSilhouette(target, ox, oy, ow, oh);
        target.strokeStyle = hexToRgba(sp3d.contourClr);
        target.lineWidth = wPx * 2;
        target.setLineDash([]);
        target.stroke();
        target.restore();
      }
    };

    // scene3d (ECMA-376 §20.1.5.5 camera projection). Active only when the
    // camera actually transforms the shape; identity/front cameras fall through
    // to the normal flat draw (but sp3d bevel/extrusion still apply via the flat
    // path below). sp3d bevel/extrusion shading is wired in just after this; the
    // contour edge is the flat approximation in strokeContourEdge.
    const scene3d = el.scene3d && isScene3dNonIdentity(el.scene3d.camera) ? el.scene3d : null;

    // Paint JUST the (clipped, optionally cropped) bitmap body at a local rect —
    // no border/contour. Separated from the edges so the bevel shading reads the
    // body silhouette's alpha without the outside-aligned contour stroke.
    const paintBitmapBodyAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      target.save();
      applyClipAt(target, ox, oy, ow, oh);
      if (crop) {
        target.drawImage(bitmap, crop.sx, crop.sy, crop.sw, crop.sh, ox, oy, ow, oh);
      } else {
        target.drawImage(bitmap, ox, oy, ow, oh);
      }
      target.restore();
    };

    // Paint the full picture body (bitmap + edges) at a local rect. Used by the
    // effect aux re-paints (reflection / soft edge) and the headless flat path,
    // which don't apply bevel. Border (a:ln) and the flat sp3d contour edge are
    // stroked AFTER the bitmap so they sit on top of / around the image.
    const paintImageAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      paintBitmapBodyAt(target, ox, oy, ow, oh);
      strokeLnBorder(target, ox, oy, ow, oh);
      strokeContourEdge(target, ox, oy, ow, oh);
    };

    // Front face for the bevel paths: the clipped bitmap PLUS its a:ln border.
    // The bevel's distance transform then starts at the border's outer edge and
    // its lit/shadowed lip shades the border itself, matching PowerPoint (and
    // the p:sp path, whose recursive body render also includes the stroke).
    const paintFaceAt = (
      target: CanvasRenderingContext2D,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      paintBitmapBodyAt(target, ox, oy, ow, oh);
      strokeLnBorder(target, ox, oy, ow, oh);
    };

    // ── sp3d bevel (ECMA-376 §20.1.5.12 bevelT/bevelB, Phase B). Device-px lip
    //    shading baked into the body bitmap before any camera projection so the
    //    rim rides the warp. The light comes from the scene's lightRig; with no
    //    rig PowerPoint's default 3-D scene light (threePt / "t") applies.
    const tf0 = ctx.getTransform();
    const det0 = Math.abs(tf0.a * tf0.d - tf0.b * tf0.c);
    const ctxDevScale = det0 > 0 ? Math.sqrt(det0) : 1;
    const bevels = buildBevelInputs(
      el.sp3d as Sp3dLike | undefined,
      el.scene3d?.lightRig as LightRigLike | undefined,
      (el.sp3d as Sp3dLike | undefined) ? (el.sp3d as { prstMaterial?: string }).prstMaterial : undefined,
      scale,
      ctxDevScale,
    );
    // Extrusion side wall only resolves under a camera that turns it visible
    // (§20.1.5.12). Computed against the active camera so a face-on view yields
    // null (no wall) and a tilted view reveals it.
    const extrusion = scene3d
      ? buildExtrusion(el.sp3d as Sp3dLike | undefined, scene3d.camera, w, h, scale, ctxDevScale)
      : null;

    // Offscreen edge margin (CSS px) for the bevel/scene3d paths: everything a
    // picture legitimately paints OUTSIDE its w×h box. Half the centre-aligned
    // border, the outside-aligned contour, the extrusion's screen sweep, plus
    // 2 px so the silhouette's antialiasing never touches the offscreen border
    // (the bevel EDT treats that border as background). See edgePadCss docs.
    const strokeHalfCss = el.stroke ? (el.stroke.width * scale) / 2 : 0;
    const contourCss = el.sp3d?.contourW ? el.sp3d.contourW * scale : 0;
    const extrusionCss = extrusion
      ? Math.hypot(extrusion.offsetX, extrusion.offsetY) / ctxDevScale
      : 0;
    const edgePadCss = strokeHalfCss + contourCss + extrusionCss + 2;

    // Draw the (clipped, optionally cropped) bitmap into a target context. This
    // is the picture "body" that the effect helpers re-paint onto aux canvases,
    // so reflections/soft edges mirror the real image rather than a flat shape.
    // When scene3d is active the body is warped through the camera homography
    // first; effects then composite over the projected result (PowerPoint
    // applies effects after the 3D transform). When sp3d carries a bevel the lip
    // shading is baked into the body before the warp (or blit, for an identity
    // camera) so it tracks the silhouette and the projection.
    const paintImage = (target: CanvasRenderingContext2D): void => {
      if (scene3d) {
        const ok = projectScene3dPaint(target, scene3d.camera, x, y, w, h, paintFaceAt, {
          bevels,
          extrusion: extrusion ?? undefined,
          paintEdges: strokeContourEdge,
          edgePadCss,
        });
        if (ok) return;
        // Headless fallback: draw flat (no bevel).
      } else if (bevels.length > 0) {
        const ok = paintBeveledFlat(
          target,
          x,
          y,
          w,
          h,
          bevels,
          paintFaceAt,
          strokeContourEdge,
          edgePadCss,
        );
        if (ok) return;
      }
      paintImageAt(target, x, y, w, h);
    };

    // Flat opaque silhouette of the clipped picture rectangle, used as the mask
    // for innerShdw / softEdge. Falls back to the bounding rect when the picture
    // is a plain rectangle (no clip path).
    const paintMaskAt = (
      target: CanvasRenderingContext2D,
      color: string,
      ox: number,
      oy: number,
      ow: number,
      oh: number,
    ): void => {
      target.save();
      applyClipAt(target, ox, oy, ow, oh);
      target.fillStyle = color;
      target.fillRect(ox, oy, ow, oh);
      target.restore();
    };

    const paintMask = (target: CanvasRenderingContext2D, color: string): void => {
      if (scene3d) {
        const ok = projectScene3dPaint(target, scene3d.camera, x, y, w, h, (octx, ox, oy, ow, oh) =>
          paintMaskAt(octx, color, ox, oy, ow, oh),
        );
        if (ok) return;
      }
      paintMaskAt(target, color, x, y, w, h);
    };

    // ── effectLst (§19.3.1.37 routes p:pic's spPr through CT_ShapeProperties,
    // so §20.1.8.16 effects apply to images). Same sequence as the p:sp path.
    const deviceW = (ctx.canvas as { width: number }).width || 0;
    const deviceH = (ctx.canvas as { height: number }).height || 0;
    const liveTransform = ctx.getTransform();
    const det = Math.abs(
      liveTransform.a * liveTransform.d - liveTransform.b * liveTransform.c,
    );
    const devScale = det > 0 ? Math.sqrt(det) : 1;
    const effBBox = { x: x * devScale, y: y * devScale, w: w * devScale, h: h * devScale };
    const effScale = scale * devScale; // EMU → device px
    const applyLiveTransform = (c: CanvasRenderingContext2D) => c.setTransform(liveTransform);
    const haveAux = deviceW > 0 && deviceH > 0;

    // Reflection sits below the picture — paint it first. §20.1.8.50. The aux
    // paint bakes in the live rotation/flip via setTransform, so the blit runs
    // at identity.
    if (el.reflection && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applyReflection(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintImage(c as CanvasRenderingContext2D); },
        effBBox, el.reflection, effScale, deviceW, deviceH,
      );
      ctx.restore();
    }

    // outerShdw / glow use the single Canvas shadow slot, cast by the image's
    // own opaque pixels. Outer shadow wins when both are present (as in p:sp).
    if (el.shadow) applyShadow(ctx, el.shadow, scale);
    else if (el.glow) applyGlow(ctx, el.glow, scale);

    // softEdge feathers the whole picture, REPLACING the direct body paint
    // (§20.1.8.53). The shadow/glow set above is carried into the aux paint via
    // setTransform of the same live context, so it still casts.
    if (el.softEdge && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applySoftEdge(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintImage(c as CanvasRenderingContext2D); },
        effBBox, el.softEdge, effScale, deviceW, deviceH,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintMask(c as CanvasRenderingContext2D, '#000'); },
      );
      ctx.restore();
    } else {
      paintImage(ctx);
    }
    if (el.shadow || el.glow) clearShadow(ctx);

    // innerShdw casts inward, on top of the picture (§20.1.8.40).
    if (el.innerShadow && haveAux) {
      ctx.save();
      ctx.setTransform(new DOMMatrix());
      applyInnerShadow(
        ctx,
        (c) => { applyLiveTransform(c as CanvasRenderingContext2D); paintMask(c as CanvasRenderingContext2D, '#000'); },
        effBBox, el.innerShadow, effScale, deviceW, deviceH,
      );
      ctx.restore();
    }

    ctx.restore();
    // bitmap is owned by getCachedBitmapByPath's cache — do not close it here.
  } catch {
    // silently skip broken images
  }
}

async function renderMedia(
  ctx: CanvasRenderingContext2D,
  el: MediaElement,
  scale: number,
  fetchMedia?: (path: string) => Promise<Blob>,
  skipControls?: boolean,
) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

  let poster: ImageBitmap | undefined;
  if (el.posterPath && fetchMedia) {
    try {
      // Poster is cached (and prefetched by renderSlide); do not close it here —
      // it is reused across renders of the same slide.
      poster = await getPosterBitmap(el, fetchMedia);
    } catch {
      // fall through to plain fill
    }
  }

  // Do not retain a saved canvas state across the asynchronous poster decode:
  // a newer render may reuse the same context while the promise is pending.
  ctx.save();
  applyFrameTransform(ctx, el, scale);
  if (poster) {
    ctx.drawImage(poster, x, y, w, h);
  } else {
    ctx.fillStyle = el.mediaKind === 'video' ? '#111' : '#f0f0f0';
    ctx.fillRect(x, y, w, h);
  }

  if (!skipControls) drawPlayBadge(ctx, x + w / 2, y + h / 2, w, h, 'paused');
  ctx.restore();
}

// ===== Table renderer =====

/**
 * Re-stroke a straight connector under one of the ECMA-376 §20.1.8.42
 * compound-line styles. Each sub-line gets its own offset along the
 * perpendicular and its own thickness; widths are taken from Office's
 * usual interpretation:
 *
 *   dbl       → two equal-thickness lines, each w/3, gap w/3
 *   thinThick → thin (w/4) + gap (w/4) + thick (w/2)
 *   thickThin → thick (w/2) + gap (w/4) + thin (w/4)
 *   tri       → three lines: thin(w/5), thick(w/5*3 ≈ 3w/5), thin(w/5)
 *
 * The starting position of each sub-line is computed so the *outer envelope*
 * stays w wide — i.e. centred on the original geometry. The base stroke
 * already drew at the centre line; we erase it (destination-out) and then
 * paint the compound sub-lines on top so dash/headEnd remain consistent.
 */
function drawCompoundLine(
  ctx: CanvasRenderingContext2D,
  start: { x: number; y: number },
  end: { x: number; y: number },
  stroke: Stroke,
  cmpd: string,
  scale: number,
): void {
  const totalW = Math.max(0.5, emuToPx(stroke.width, scale));
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const len = Math.hypot(dx, dy);
  if (len === 0) return;
  const px = -dy / len;
  const py = dx / len;

  // sub: relative position along the perpendicular axis (-1..1, with 0 being
  // the centre line) and width as a fraction of totalW.
  type Sub = { offset: number; widthFrac: number };
  let subs: Sub[];
  switch (cmpd) {
    case 'dbl':
      subs = [
        { offset: -1 / 3, widthFrac: 1 / 3 },
        { offset:  1 / 3, widthFrac: 1 / 3 },
      ];
      break;
    case 'thinThick':
      subs = [
        { offset: -3 / 8, widthFrac: 1 / 4 },
        { offset:  1 / 4, widthFrac: 1 / 2 },
      ];
      break;
    case 'thickThin':
      subs = [
        { offset: -1 / 4, widthFrac: 1 / 2 },
        { offset:  3 / 8, widthFrac: 1 / 4 },
      ];
      break;
    case 'tri':
      subs = [
        { offset: -2 / 5, widthFrac: 1 / 5 },
        { offset:  0,     widthFrac: 3 / 5 },
        { offset:  2 / 5, widthFrac: 1 / 5 },
      ];
      break;
    default:
      return;
  }

  ctx.save();
  // 1. Erase the centre line that was already drawn at full width.
  ctx.globalCompositeOperation = 'destination-out';
  ctx.strokeStyle = '#000';
  ctx.lineWidth = totalW + 0.5; // small overshoot to fully erase antialiasing fringe
  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(start.x, start.y); ctx.lineTo(end.x, end.y);
  ctx.stroke();

  // 2. Paint each sub-line.
  ctx.globalCompositeOperation = 'source-over';
  ctx.strokeStyle = hexToRgba(stroke.color);
  for (const sub of subs) {
    const ox = px * (totalW * sub.offset);
    const oy = py * (totalW * sub.offset);
    ctx.lineWidth = Math.max(0.5, totalW * sub.widthFrac);
    ctx.beginPath();
    ctx.moveTo(start.x + ox, start.y + oy);
    ctx.lineTo(end.x + ox, end.y + oy);
    ctx.stroke();
  }
  ctx.restore();
}

function applyStroke(ctx: CanvasRenderingContext2D, stroke: Stroke | null, scale: number) {
  // `scale` is EMU → px factor (canvasWidthPx / slideWidthEMU).
  applyStrokeCore(ctx, stroke, scale);
}

// ─── Chart rendering ────────────────────────────────────────────────────────
// Chart rendering is delegated to @silurus/ooxml-core's unified renderer.
// See renderChart(ctx, el, scale) below.


// ─── Table rendering ─────────────────────────────────────────────────────────

export function applyFrameTransform(
  ctx: CanvasRenderingContext2D,
  el: Pick<TableElement | ChartElement | MediaElement, 'x' | 'y' | 'width' | 'height' | 'rotation' | 'flipH' | 'flipV'>,
  scale: number,
): void {
  if (el.rotation === 0 && !el.flipH && !el.flipV) return;
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);
  ctx.translate(x + w / 2, y + h / 2);
  ctx.rotate((el.rotation * Math.PI) / 180);
  if (el.flipH) ctx.scale(-1, 1);
  if (el.flipV) ctx.scale(1, -1);
  ctx.translate(-(x + w / 2), -(y + h / 2));
}

export function renderTable(ctx: CanvasRenderingContext2D, el: TableElement, scale: number, slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 }) {
  ctx.save();
  applyFrameTransform(ctx, el, scale);
  const x0 = emuToPx(el.x, scale);
  const y0 = emuToPx(el.y, scale);

  // Convert col widths to pixels.
  const colWidths = el.cols.map(c => emuToPx(c, scale));
  const numCols = colWidths.length;

  // Spanned width starting at column `ci` for `span` columns.
  const spannedWidth = (ci: number, span: number): number => {
    let w = 0;
    for (let s = 0; s < span; s++) w += colWidths[ci + s] ?? 0;
    return w;
  };

  // ── Row heights: ECMA-376 §21.1.3.18 (a:tr@h) is a MINIMUM ────────────────
  // PowerPoint grows a row to fit its tallest cell's laid-out text (like
  // Word's "at least" line rule). A literal h=0 therefore becomes
  // content-driven. We measure each cell's text body at its spanned width
  // (reusing the same renderTextBody machinery via measureOnly) and take
  // max(tr@h, tallest single-row cell content). A rowSpan cell distributes
  // its content height across the rows it covers so it doesn't inflate the
  // first row.
  const rowHeights = el.rows.map(r => emuToPx(r.height, scale));

  // First pass: single-row (rowSpan ≤ 1) cells set their own row's minimum.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      if ((cell.rowSpan || 1) > 1) continue;
      if (!cell.textBody) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const needed = (renderTextBody(
        ctx, cell.textBody, 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true,
      ) as number) || 0;
      if (needed > rowHeights[ri]) rowHeights[ri] = needed;
    }
  }

  // Second pass: rowSpan cells. If the content needs more than the sum of the
  // rows it spans, distribute the deficit across those rows so the merged area
  // grows without inflating any single row beyond what its own content needs.
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      if (cell.hMerge || cell.vMerge) continue;
      const span = cell.rowSpan || 1;
      if (span <= 1 || !cell.textBody) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      const needed = (renderTextBody(
        ctx, cell.textBody, 0, 0, cellW, 0, scale, null, 0, false, false,
        '#000000', slideNumber, rc, undefined, true,
      ) as number) || 0;
      let have = 0;
      for (let s = 0; s < span && ri + s < rowHeights.length; s++) have += rowHeights[ri + s];
      if (needed > have) {
        const extra = (needed - have) / span;
        for (let s = 0; s < span && ri + s < rowHeights.length; s++) rowHeights[ri + s] += extra;
      }
    }
  }

  // ── Column x-positions ────────────────────────────────────────────────────
  // ECMA-376 §21.1.3.13 (a:tblPr@rtl): a right-to-left table places column 0
  // at the right edge, columns advancing leftward. We precompute the left
  // pixel edge of each column so RTL is a coordinate flip (no ctx.scale(-1,1)
  // which would mirror text and borders).
  const tableW = colWidths.reduce((a, b) => a + b, 0);
  const colLeft: number[] = new Array(numCols);
  if (el.rtl) {
    let right = x0 + tableW;
    for (let ci = 0; ci < numCols; ci++) {
      right -= colWidths[ci];
      colLeft[ci] = right;
    }
  } else {
    let left = x0;
    for (let ci = 0; ci < numCols; ci++) {
      colLeft[ci] = left;
      left += colWidths[ci];
    }
  }

  // Left pixel edge of a cell spanning columns [ci, ci+span). Under RTL the
  // span grows leftward, so the merged cell's left edge is the leftmost column.
  const spannedLeft = (ci: number, span: number): number =>
    el.rtl ? colLeft[ci + span - 1] : colLeft[ci];

  const rowTop: number[] = new Array(el.rows.length);
  {
    let y = y0;
    for (let ri = 0; ri < el.rows.length; ri++) { rowTop[ri] = y; y += rowHeights[ri]; }
  }

  // ECMA-376 border-collapse: paint every cell's fill + text body FIRST, then
  // every cell's borders, so a neighbouring cell's background fill can never
  // overpaint the half of a shared gridline an adjacent cell already drew.
  // (Mirrors the docx table renderer's two-pass order.) Collect each
  // non-continuation cell's paint box once; rowSpan/gridSpan and the RTL
  // coordinate flip are already baked into colX/rowY/cellW/cellH, so both passes
  // replay identical geometry.
  // Each job also records its GRID footprint (top-left slot ci/ri, its column
  // span and last row) so the border pass can find, for every interior gridline,
  // the neighbouring cell that shares it — DrawingML tables populate every grid
  // slot (a spanning `<a:tc>` is followed by `hMerge`/`vMerge` continuation
  // `<a:tc>`s), so `ci` is the true grid column index.
  const jobs: Array<{
    cell: TableCell; colX: number; rowY: number; cellW: number; cellH: number;
    ci: number; ri: number; span: number; lastRi: number;
  }> = [];
  // Grid-slot → job index occupancy, so an interior edge can look up the adjacent
  // cell. A spanning anchor fills every slot it covers; the neighbour query on any
  // covered slot resolves back to that anchor job.
  const occupancy: number[][] = el.rows.map(() => new Array<number>(numCols).fill(-1));
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    const rowY = rowTop[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      // Merged-cell continuations are drawn by their anchor cell.
      if (cell.hMerge || cell.vMerge) continue;
      const span = cell.gridSpan || 1;
      const rspan = cell.rowSpan || 1;
      const cellW = spannedWidth(ci, span);
      let cellH = 0;
      for (let s = 0; s < rspan; s++) cellH += rowHeights[ri + s] ?? 0;
      const colX = spannedLeft(ci, span);
      const lastRi = Math.min(ri + rspan - 1, el.rows.length - 1);
      const jobIndex = jobs.length;
      jobs.push({ cell, colX, rowY, cellW, cellH, ci, ri, span, lastRi });
      // Record this cell's grid footprint so interior-edge neighbour lookups
      // resolve to it (a spanning cell owns every slot it covers).
      for (let rj = ri; rj <= lastRi; rj++) {
        for (let cj = ci; cj < ci + span && cj < numCols; cj++) {
          occupancy[rj][cj] = jobIndex;
        }
      }
    }
  }

  // Pass 1: fills + text bodies.
  for (const { cell, colX, rowY, cellW, cellH } of jobs) {
    const fillColor = resolveFill(cell.fill);
    if (fillColor) {
      ctx.fillStyle = fillColor;
      ctx.fillRect(colX, rowY, cellW, cellH);
    }
    // Text body — default run colour comes from the table style's tcTxStyle
    // (e.g. white header text on an accent fill); a run's explicit colour wins.
    if (cell.textBody) {
      const cellDefaultColor = cell.textColor ? hexToRgba(cell.textColor) : null;
      renderTextBody(ctx, cell.textBody, colX, rowY, cellW, cellH, scale, cellDefaultColor, 0, false, false, '#000000', slideNumber, rc);
    }
  }

  // Pass 2: borders, painted on top of every fill.
  //
  // SPEC IS SILENT on adjacent-cell border conflict for DrawingML/PresentationML
  // tables (unlike WordprocessingML §17.4.66). When cell spacing is zero, two
  // neighbouring cells each contribute a line for the SHARED interior gridline;
  // the OLD code drew both, so the later-painted cell won by paint order (and a
  // translucent line doubled its ink). We now draw each gridline EXACTLY ONCE via
  // a single-ownership convention (the pptx leg / structural mirror of docx
  // PR #811, §17.4.66):
  //   • outer table edges → the single bordering cell draws its own line;
  //   • interior VERTICAL line → the physically-LEFT cell owns it (drawn as its
  //     physical-right edge), resolved against the physically-right neighbour;
  //   • interior HORIZONTAL line → the physically-ABOVE cell owns it (drawn as its
  //     bottom edge), resolved against the below neighbour.
  // The winner between two conflicting facing lines is chosen by our DEFINED,
  // deterministic rule (spec silent): null loses → wider wins → darker wins →
  // owner (reading-order-first) wins. See `resolveTableBorderConflict`.
  //
  // Crispness nudge (see crispOffset): an axis-aligned thin grid line whose
  // device-pixel width is odd straddles two device rows/cols on a DPR=1 display
  // (each ~50% ink → blurry). Snapping the cell edge perpendicular to the line
  // onto the nearest crisp device position centers an odd-width stroke on one
  // device row/col → crisp. `applyStroke` sets ctx.lineWidth to the logical
  // width, so we read it back to pick the right offset, and the snap delta is
  // derived from the edge's own coordinate (fractional-safe). Horizontal edges
  // (T/B) shift Y from their y; vertical edges (L/R) shift X from their x.
  // Diagonals can't be pixel-aligned, so untouched.
  const dpr = rc.dpr;

  // Resolve the job whose grid footprint covers slot (ri, ci), or null when the
  // slot is empty / out of range (an occupancy value of -1 or an index off-grid).
  const jobAt = (ri: number, ci: number): typeof jobs[number] | null => {
    if (ri < 0 || ri >= occupancy.length) return null;
    if (ci < 0 || ci >= numCols) return null;
    const idx = occupancy[ri][ci];
    return idx < 0 ? null : jobs[idx];
  };

  const strokeSeg = (
    stroke: Stroke,
    x1: number, y1: number, x2: number, y2: number,
  ) => {
    applyStroke(ctx, stroke, scale);
    // Vertical edge (x1===x2) nudges X; horizontal edge (y1===y2) nudges Y.
    const dx = x1 === x2 ? crispOffset(x1, ctx.lineWidth, dpr) : 0;
    const dy = y1 === y2 ? crispOffset(y1, ctx.lineWidth, dpr) : 0;
    ctx.beginPath();
    ctx.moveTo(x1 + dx, y1 + dy);
    ctx.lineTo(x2 + dx, y2 + dy);
    ctx.stroke();
  };

  for (const j of jobs) {
    const { cell, colX, rowY, cellW, cellH } = j;
    ctx.save();

    // Map this cell's LOGICAL border specs onto PHYSICAL sides. Under a
    // right-to-left table (§21.1.3.13) logical column 0 is at the physical RIGHT,
    // so a cell's logical-left border (`borderL`) paints on its physical-right
    // edge and its logical-right border (`borderR`) on its physical-left edge;
    // the physically-right NEIGHBOUR sits at a LOWER logical column. LTR keeps the
    // natural mapping.
    const physLeftSpec = el.rtl ? cell.borderR : cell.borderL;
    const physRightSpec = el.rtl ? cell.borderL : cell.borderR;
    const physLeftOuter = el.rtl ? j.ci + j.span === numCols : j.ci === 0;
    const physRightOuter = el.rtl ? j.ci === 0 : j.ci + j.span === numCols;
    // Grid slot of the physically-right neighbour (whose facing edge conflicts
    // with THIS cell's physical-right line). LTR: the column just past this span.
    // RTL: the column just before this cell's leading logical column.
    const physRightNbrCi = el.rtl ? j.ci - 1 : j.ci + j.span;
    // The neighbour's facing edge is the one on ITS physical-left side, i.e. its
    // logical-right spec under RTL, logical-left spec otherwise.
    const nbrPhysLeftSpec = (nb: TableCell): Stroke | null => (el.rtl ? nb.borderR : nb.borderL);

    // TOP: only the outer top row draws its own top; interior tops are owned by
    // the cell above (its bottom). Vertical direction is unaffected by RTL.
    if (j.ri === 0 && cell.borderT) {
      strokeSeg(cell.borderT, colX, rowY, colX + cellW, rowY);
    }
    // PHYSICAL LEFT: only the outer-left column draws its own left; interior
    // lefts are owned by the physically-left neighbour (its right edge).
    if (physLeftOuter && physLeftSpec) {
      strokeSeg(physLeftSpec, colX, rowY, colX, rowY + cellH);
    }
    // BOTTOM: outer bottom → own spec; interior → resolve THIS cell's bottom
    // against the below neighbour's top and draw the single winner.
    {
      if (j.lastRi === el.rows.length - 1) {
        const spec = cell.borderB;
        if (spec) strokeSeg(spec, colX, rowY + cellH, colX + cellW, rowY + cellH);
      } else {
        // #824 — the shared horizontal edge below this cell may face SEVERAL finer
        // below-cells with differing borderT (this cell is wider via gridSpan).
        // Subdivide the edge at those cells' column boundaries and resolve EACH
        // sub-segment against its OWN below neighbour, drawing a per-segment winner
        // rather than resolving the whole span against the origin slot alone.
        const belowRi = j.lastRi + 1;
        const y = rowY + cellH;
        const cEndMax = Math.min(j.ci + j.span, numCols);
        let cj = j.ci;
        while (cj < cEndMax) {
          const idx = occupancy[belowRi][cj];
          let cEnd = cj + 1;
          while (cEnd < cEndMax && occupancy[belowRi][cEnd] === idx) cEnd++;
          const below = jobAt(belowRi, cj);
          const spec = resolveTableBorderConflict(cell.borderB, below ? below.cell.borderT : null);
          if (spec) {
            const segLeft = spannedLeft(cj, cEnd - cj);
            const segW = spannedWidth(cj, cEnd - cj);
            strokeSeg(spec, segLeft, y, segLeft + segW, y);
          }
          cj = cEnd;
        }
      }
    }
    // PHYSICAL RIGHT: outer-right → own spec; interior → resolve THIS cell's
    // physical-right line against the physically-right neighbour's facing edge
    // (so each shared vertical line is drawn once).
    {
      if (physRightOuter) {
        const spec = physRightSpec;
        if (spec) strokeSeg(spec, colX + cellW, rowY, colX + cellW, rowY + cellH);
      } else {
        // #824 — a rowSpan cell's physical-right edge may face SEVERAL finer
        // right-neighbours down the rows it spans, with differing facing specs.
        // Subdivide the edge at those neighbours' row boundaries and resolve EACH
        // sub-segment against its OWN neighbour's facing (physical-left) edge.
        const x = colX + cellW;
        let rj = j.ri;
        while (rj <= j.lastRi) {
          const idx = occupancy[rj][physRightNbrCi];
          let rEnd = rj;
          while (rEnd + 1 <= j.lastRi && occupancy[rEnd + 1][physRightNbrCi] === idx) rEnd++;
          const right = jobAt(rj, physRightNbrCi);
          const spec = resolveTableBorderConflict(physRightSpec, right ? nbrPhysLeftSpec(right.cell) : null);
          if (spec) strokeSeg(spec, x, rowTop[rj], x, rowTop[rEnd] + rowHeights[rEnd]);
          rj = rEnd + 1;
        }
      }
    }

    // Diagonal borders: top-left→bottom-right and bottom-left→top-right. These
    // are never shared between cells, so they are always drawn per-cell (and are
    // not pixel-aligned).
    if (cell.diagonalTL) {
      applyStroke(ctx, cell.diagonalTL, scale);
      ctx.beginPath();
      ctx.moveTo(colX, rowY);
      ctx.lineTo(colX + cellW, rowY + cellH);
      ctx.stroke();
    }
    if (cell.diagonalTR) {
      applyStroke(ctx, cell.diagonalTR, scale);
      ctx.beginPath();
      ctx.moveTo(colX + cellW, rowY);
      ctx.lineTo(colX, rowY + cellH);
      ctx.stroke();
    }
    ctx.restore();
  }
  ctx.restore();
}

// ===== Public API =====

export type { RenderOptions } from './types';

/**
 * Overlay a translucent rectangle over the whole slide so a hidden slide reads
 * faintly (PowerPoint's hidden-slide thumbnail look). A pure mechanism — it
 * never inspects whether the slide is hidden; the caller ({@link PptxViewer}'s
 * `'dim'` mode) decides when to apply it. `widthPx`/`heightPx` are CSS px; the
 * ctx is already `scale(dpr, dpr)`d, so a fillRect in CSS px covers the canvas.
 */
export function applyDimOverlay(
  ctx: CanvasRenderingContext2D,
  dim: DimOptions,
  widthPx: number,
  heightPx: number,
): void {
  ctx.save();
  ctx.globalAlpha = dim.opacity;
  ctx.fillStyle = dim.color;
  ctx.fillRect(0, 0, widthPx, heightPx);
  ctx.restore();
}

/**
 * Internal render options: the shared {@link RenderOptions} plus the opt-in
 * `math` engine. `math` is internal plumbing — the headless {@link
 * PptxPresentation} injects it once at load and threads it here on each draw,
 * so the public `RenderSlideOptions` deliberately does not expose it.
 */
type SlideRenderOptions = RenderOptions & { math?: MathRenderer; dim?: DimOptions };

/**
 * Per-canvas monotonic render token for the {@link renderSlide} cancellation
 * guard. A WeakMap keyed on the canvas replaces the previous property monkey-
 * patch (`canvas.__pptxRenderToken`), so no non-standard field is written onto
 * the caller's canvas and the `as unknown as` cast is gone. WeakMap keys are
 * held weakly, so a discarded canvas is collected normally.
 */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

/**
 * RB7: paint a placeholder for a slide whose part failed to parse. A neutral
 * card with a warning glyph, a heading, and the part-tagged error — so a viewer
 * shows "this one slide is broken" in place, and the rest of the deck renders
 * normally. Widths are in CSS px (the ctx is already dpr-scaled). Only ever
 * called for a slide carrying `parseError`, so healthy slides never touch it.
 */
function drawParseErrorPlaceholder(
  ctx: CanvasRenderingContext2D,
  widthPx: number,
  heightPx: number,
  slideNumber: number,
  message: string,
): void {
  ctx.save();
  // Neutral card fill + dashed border, inset from the slide edge.
  ctx.fillStyle = '#f7f7f8';
  ctx.fillRect(0, 0, widthPx, heightPx);
  const pad = Math.max(12, Math.min(widthPx, heightPx) * 0.04);
  ctx.strokeStyle = '#c8ccd2';
  ctx.lineWidth = Math.max(1, Math.min(widthPx, heightPx) * 0.004);
  ctx.setLineDash([ctx.lineWidth * 6, ctx.lineWidth * 5]);
  ctx.strokeRect(pad, pad, widthPx - pad * 2, heightPx - pad * 2);
  ctx.setLineDash([]);

  const cx = widthPx / 2;
  // Warning glyph, scaled to the card.
  const glyph = Math.max(18, Math.min(widthPx, heightPx) * 0.14);
  ctx.fillStyle = '#b23b3b';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.font = `${glyph}px sans-serif`;
  ctx.fillText('⚠', cx, heightPx * 0.34);

  // Heading.
  const headSize = Math.max(11, Math.min(widthPx, heightPx) * 0.045);
  ctx.fillStyle = '#333333';
  ctx.font = `600 ${headSize}px sans-serif`;
  ctx.fillText(`Slide ${slideNumber} could not be displayed`, cx, heightPx * 0.52);

  // Error detail (part path + reason), wrapped to the card width and clipped to
  // a few lines so a long message never overflows the frame.
  const detailSize = Math.max(9, Math.min(widthPx, heightPx) * 0.028);
  ctx.fillStyle = '#666666';
  ctx.font = `${detailSize}px sans-serif`;
  const maxLineWidth = widthPx - pad * 4;
  const words = message.split(/\s+/);
  const lines: string[] = [];
  let line = '';
  for (const word of words) {
    const candidate = line ? `${line} ${word}` : word;
    if (ctx.measureText(candidate).width > maxLineWidth && line) {
      lines.push(line);
      line = word;
    } else {
      line = candidate;
    }
    if (lines.length >= 4) break; // cap so it never overruns the card
  }
  if (line && lines.length < 4) lines.push(line);
  const lineHeight = detailSize * 1.35;
  let y = heightPx * 0.6 + lineHeight;
  for (const l of lines.slice(0, 4)) {
    ctx.fillText(l, cx, y);
    y += lineHeight;
  }
  ctx.restore();
}

/**
 * Render a single slide onto a <canvas> element.
 * Returns the canvas for convenience.
 */
export async function renderSlide(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: SlideRenderOptions = {},
  onTextRun?: TextRunCallback,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  // Render-pass lease (core acquireBitmapCacheLease): the warm pass below fires
  // a decode for every picture on the slide and the draw loop then awaits each
  // element's bitmap and draws it. The shared bitmap cache is LRU-bounded, so a
  // slide referencing more images than the cap — or a concurrent render on the
  // same deck — could evict AND GPU-close a bitmap between the draw loop's await
  // and its drawImage. Under the lease the eviction still removes the cache
  // entry (bounded size; a later resolve re-decodes), but the close is deferred
  // until this pass ends, so no draw ever receives a closed bitmap.
  const releaseLease = opts.fetchImage ? acquireBitmapCacheLease(opts.fetchImage) : undefined;
  try {
    return await renderSlideLeased(canvas, slide, slideWidth, slideHeight, opts, onTextRun);
  } finally {
    releaseLease?.();
  }
}

/** {@link renderSlide}'s body, verbatim; runs under the caller's render-pass
 *  lease. */
async function renderSlideLeased(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  slide: Slide,
  slideWidth: number,
  slideHeight: number,
  opts: SlideRenderOptions = {},
  onTextRun?: TextRunCallback,
): Promise<HTMLCanvasElement | OffscreenCanvas> {
  // Cancellation guard. renderSlide is async (it awaits image / equation decode),
  // so rapid navigation can start a newer render of the SAME canvas before this
  // one finishes. Both would `canvas.width = …` (clear) and then draw, and their
  // draws interleave at the await points — ghosting multiple slides together.
  // Stamp a per-canvas token; once a newer render supersedes us, stop drawing at
  // the next await so only the latest render's output survives.
  const myToken = (renderTokens.get(canvas) ?? 0) + 1;
  renderTokens.set(canvas, myToken);
  const superseded = () => renderTokens.get(canvas) !== myToken;

  const targetWidth = opts.width ?? ((isHTMLCanvas(canvas) ? canvas.offsetWidth : 0) || 960);
  const scale = targetWidth / slideWidth;
  const canvasW = Math.round(targetWidth);
  const canvasH = Math.round(slideHeight * scale);

  const dpr = opts.dpr ?? defaultDpr();
  // Clamp the backing store to browser canvas limits (RB5). A huge target width
  // (or large dpr × slide size) can exceed the per-axis / total-area cap, at
  // which point the browser silently allocates a smaller-or-empty buffer and the
  // slide renders blank. `clampCanvasSize` scales BOTH axes by one factor (≤ 1)
  // so the aspect ratio is kept; we fold that into the effective dpr, keep the
  // CSS box at its intended size, and the browser stretches the (slightly
  // lower-res) backing store to fill it — a visible slide beats a blank one.
  const clamped = clampCanvasSize(canvasW * dpr, canvasH * dpr);
  const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;
  canvas.width = clamped.width;
  canvas.height = clamped.height;
  // CSS size only applies to the visible HTMLCanvasElement (not OffscreenCanvas)
  if (isHTMLCanvas(canvas)) {
    canvas.style.width = `${canvasW}px`;
    // Mirror the docx renderer: when callers use `renderSlide(canvas, ...)`
    // directly without the {@link PptxViewer} wrapper, set `display:block`
    // as a safety net so the inline-element baseline does not leave a
    // descender gap below the canvas. Respect any user-specified value.
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) throw new Error('Could not get 2D context');
  // Use the effective dpr (folded with any clamp factor) so drawing fills the
  // clamped backing store and crisp-offset math stays aligned with it.
  ctx.scale(effectiveDpr, effectiveDpr);

  // RB7 partial degradation: a slide whose part failed to parse (see the Rust
  // `broken_slide`) carries `parseError` and no elements. Paint a visible error
  // placeholder — correctly sized so navigation geometry is unchanged — instead
  // of a blank frame, and stop. Healthy slides (no parseError) are unaffected.
  if (slide.parseError) {
    drawParseErrorPlaceholder(ctx, canvasW, canvasH, slide.slideNumber, slide.parseError);
    return canvas;
  }

  const themeDefaultColor = opts.defaultTextColor
    ? `#${opts.defaultTextColor}`
    : '#000000';

  const rc: RenderContext = {
    themeMajorFont: opts.majorFont ?? null,
    themeMinorFont: opts.minorFont ?? null,
    themeHlinkColor: opts.hlinkColor ?? null,
    // The backing store may have been clamped below `canvasSize × dpr`; downstream
    // crisp-offset math must use the SAME effective dpr the ctx was scaled by.
    dpr: effectiveDpr,
    // Issue #805 — legible default for the synthetic SmartArt fallback shape's
    // null-colour runs, derived once per slide from the background luminance.
    smartArtFallbackTextColor: smartArtFallbackTextColor(slide.background, themeDefaultColor),
  };

  await renderBackground(ctx, slide.background, canvasW, canvasH, scale, opts.fetchImage);
  if (superseded()) return canvas;

  // Pre-rasterize any equations so the synchronous text layout can place them.
  // `math` is the engine injected once at PptxPresentation.load and threaded in
  // here; without it, equations are skipped and the asset never enters the bundle.
  if (opts.math) await prepareSlideMath(slide, opts.math);
  if (superseded()) return canvas;

  const slideNumber = slide.slideNumber;

  // Warm the bitmap caches for every image-bearing element concurrently.
  // The draw loop below still awaits in element order (z-order), but each
  // await now hits a settled/in-flight promise instead of starting a serial
  // fetch+decode — first paint cost becomes max(decode) instead of sum.
  for (const el of slide.elements) {
    if (el.type === 'picture' && opts.fetchImage) {
      // Warm exactly the source the draw loop (renderPicture) will await: the
      // SVG decode when the picture carries an svgImagePath and no srcRect crop,
      // otherwise the raster bitmap. This mirrors the draw-path source selection
      // above, so the await there hits a settled/in-flight promise instead of
      // starting a serial fetch + decode. Warming the raster for an uncropped
      // SVG-bearing picture would instead leave the hot (SVG) cache cold and
      // waste a fetch + createImageBitmap on a fallback that is never drawn.
      // (The draw path still falls back to getCachedBitmapByPath on SVG decode
      // failure, so the raster stays cold only in that rare case.)
      const p = el as PictureElement;
      const pDataIsSvg = p.mimeType === 'image/svg+xml';
      if (preferVectorBlip(p)) {
        void getCachedSvgImageByPath(p.svgImagePath, opts.fetchImage).catch(() => undefined);
      } else if (pDataIsSvg) {
        void getCachedSvgImageByPath(p.imagePath, opts.fetchImage).catch(() => undefined);
      } else {
        // Pass the picture's pt size so a metafile blip warms at the same raster
        // size the draw loop requests (the cache is path-keyed, first-wins). A
        // cropped metafile warms at its full picture frame, matching the draw
        // path's `metafileRasterSize` call.
        const warm = metafileRasterSize(
          p.mimeType,
          p.srcRect,
          p.width / PT_TO_EMU,
          p.height / PT_TO_EMU,
        );
        // Warm through the duotone cache so a §20.1.8.23 recolour picture warms
        // its recoloured variant (keyed by path + colours); no duotone ⇒ this is
        // the plain base-bitmap warm, byte-identical to before.
        void getCachedDuotoneBitmapByPath(p.imagePath, p.mimeType, p.duotone, opts.fetchImage, {
          widthPt: warm.widthPt,
          heightPt: warm.heightPt,
        }).catch(() => undefined);
      }
    } else if (el.type === 'media') {
      const m = el as MediaElement;
      if (m.posterPath && opts.fetchMedia) {
        void getPosterBitmap(m, opts.fetchMedia).catch(() => undefined);
      }
    }
  }

  // Picture bullets (`<a:buBlip>`, §21.1.2.4.2) are drawn inside the SYNCHRONOUS
  // text-body layout, which can't await a decode. Resolve every bullet image up
  // front (deduped by path via getCachedBitmapByPath) and await them so the draw
  // loop's peekCachedBitmapByPath finds a settled bitmap. Missing/failed decodes resolve to
  // undefined and the marker is simply skipped — never blocking the frame.
  if (opts.fetchImage) {
    const fetchImage = opts.fetchImage;
    const bulletPaths = new Set<string>();
    for (const el of slide.elements) {
      if (el.type !== 'shape' || !el.textBody) continue;
      for (const para of el.textBody.paragraphs) {
        const b = asBullet(para.bullet);
        if (b.type === 'blip') bulletPaths.add(`${b.imagePath} ${b.mimeType}`);
      }
    }
    if (bulletPaths.size > 0) {
      await Promise.all(
        [...bulletPaths].map((key) => {
          const [path, mime] = key.split(' ');
          return getCachedBitmapByPath(path, mime, fetchImage).catch(() => undefined);
        }),
      );
      if (superseded()) return canvas;
    }
  }

  for (const el of slide.elements) {
    // A newer render of this canvas started while we awaited an image/equation —
    // stop so we don't paint this (now stale) slide over the newer one.
    if (superseded()) return canvas;
    if (el.type === 'shape') {
      renderShape(ctx, el, scale, themeDefaultColor, slideNumber, rc, onTextRun, opts.fetchImage);
    } else if (el.type === 'picture') {
      await renderPicture(ctx, el, scale, opts.fetchImage);
    } else if (el.type === 'table') {
      renderTable(ctx, el, scale, slideNumber, rc);
    } else if (el.type === 'media') {
      await renderMedia(ctx, el, scale, opts.fetchMedia, opts.skipMediaControls);
    } else if (el.type === 'chart') {
      // OOXML: 1pt = 12700 EMU. The slide renderer's `scale` is px-per-EMU,
      // so PT_TO_EMU * scale gives pixels-per-point at the current display size.
      const chartPtToPx = PT_TO_EMU * scale;
      // `el.chart` is already the canonical ChartModel emitted by the Rust
      // parser (`ooxml_common::chart::ChartModel`) — no per-field adapter.
      ctx.save();
      applyFrameTransform(ctx, el, scale);
      renderChart(
        ctx,
        el.chart,
        {
          x: emuToPx(el.x, scale),
          y: emuToPx(el.y, scale),
          w: emuToPx(el.width, scale),
          h: emuToPx(el.height, scale),
        },
        chartPtToPx,
      );
      ctx.restore();
    }
  }

  // Hidden-slide dimming (a render mechanism; the caller decides WHEN — see
  // PptxViewer's 'dim' mode). Overlay AFTER all content so it reads faintly.
  if (superseded()) return canvas;
  if (opts.dim) applyDimOverlay(ctx, opts.dim, canvasW, canvasH);

  return canvas;
}
