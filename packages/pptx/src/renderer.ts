import type {
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  MediaElement,
  TableElement,
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
  TabStop,
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
  classifyCjkFont,
  classifyFontGeneric,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  DEFAULT_KINSOKU_RULES,
  isCjkBreakChar,
  getCachedSvgImageByPath,
  getCachedBitmapByPath,
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
  measureAdvanceWithTrailingSpace,
} from '@silurus/ooxml-core';
import type { CameraInput, Vec2, BevelInput, ExtrusionInput, BevelRegion } from '@silurus/ooxml-core';
import type { MathNode, MathRenderer } from '@silurus/ooxml-core';
import { drawPlayBadge } from './media-chrome';
import {
  segmentsHaveRtl,
  computeLineVisualOrder,
  type LineVisualOrder,
} from './bidi-line';
import { fitCjkLine, type MeasuredChar } from './cjk-wrap.js';
import { justifyLine, type Justified } from './text-justify';
import { justifiedPiecePositions } from '@silurus/ooxml-core';

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
}

/** Information about a rendered text segment for building a transparent selection overlay. */
export interface PptxTextRunInfo {
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
  /** Segments right-aligned at a tab stop (set when paragraph contains \t and a right-aligned tabStop) */
  tabStop?: {
    /** Tab stop position in px from the left edge of the text area (bx + lPad + tabStop.px = canvas X) */
    px: number;
    algn: string;
    segments: LayoutSegment[];
  };
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

  // Tab stop state: once we hit a \t we switch to collecting tabStop.segments
  let tabActive = false;
  let tabStopPx = 0;   // position of tab stop from text area left (px)

  const newLine = (endsWithBreak = false) => {
    if (endsWithBreak) currentLine.endsWithBreak = true;
    lines.push(currentLine);
    currentLine = { segments: [] };
    lineW = 0;
    tabActive = false; // reset tab state per line
    hasWhitespaceOnLine = false;
  };

  // Push to the active segment list (main or tab-stop group)
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
    },
  ) => {
    if (!text) return;
    ctx.font = font;
    const lsPx = extras?.letterSpacingPx ?? 0;
    // `measureAdvanceWithTrailingSpace` (not raw measureText): runs are tokenized
    // on whitespace (`split(/(\s+)/)`), so an inter-word space is pushed as a lone
    // " " token whose `measureText(" ").width` is 0 on trimming engines (Firefox,
    // WebKit/Safari) — the token would carry no advance and adjacent words would
    // collide. The core helper restores the space advance; on Chrome/skia it is a
    // pure pass-through. The draw pen advances by this same stored `w` (no
    // re-measure), so measure==draw on every engine.
    const baseW = measureAdvanceWithTrailingSpace(ctx, text);
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
    // Shadow / outline use object identity for merging — adjacent runs share
    // the same object since the run is parsed once. Different objects (or
    // one set / one missing) force a new segment.
    const sameMeta = (a: LayoutSegment) =>
      !a.math &&
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
      (a.fontFamily ?? '') === (fontFamily ?? '');
    if (tabActive && currentLine.tabStop) {
      const segs = currentLine.tabStop.segments;
      const last = segs.at(-1);
      if (last && sameMeta(last)) {
        last.text += text;
      } else {
        segs.push({ text, font, fontFamily, sizePx, color, underline, underlineStyle, underlineColor, strikethrough, strikeDouble, letterSpacingPx: lsPx || undefined, baseline, shadow, outline, highlight });
      }
    } else {
      lineW += w;
      const last = currentLine.segments.at(-1);
      if (last && sameMeta(last)) {
        last.text += text;
      } else {
        currentLine.segments.push({ text, font, fontFamily, sizePx, color, underline, underlineStyle, underlineColor, strikethrough, strikeDouble, letterSpacingPx: lsPx || undefined, baseline, shadow, outline, highlight });
      }
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
      else if (lineW + width > lineMaxW() && lineW > 0) newLine();
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
    };

    // Split on whitespace boundaries, keeping the whitespace tokens
    const tokens = runText.split(/(\s+)/);

    for (const token of tokens) {
      if (!token) continue;

      // ── Tab character ────────────────────────────────────────────────────
      if (/^\t+$/.test(token)) {
        // Find first tab stop whose position (from text area left) is beyond the current pen
        const currentAbsW = marLPx + lineW; // current position from text area left
        const ts = (para.tabStops ?? []).find(
          (t: TabStop) => emuToPx(t.pos, scale) > currentAbsW
        );
        if (ts) {
          tabStopPx = emuToPx(ts.pos, scale);
          if (ts.algn === 'r' || ts.algn === 'ctr') {
            // Switch to tab-stop accumulation mode
            tabActive = true;
            currentLine.tabStop = { px: tabStopPx, algn: ts.algn, segments: [] };
          } else {
            // Left-aligned tab: advance lineW to the tab stop
            lineW = tabStopPx - marLPx;
          }
        } else {
          // No matching tab stop — treat as a single space
          push(' ', font, sizePx, color, segUnderline, run.strikethrough, undefined, segExtras);
        }
        continue;
      }

      ctx.font = font;
      // Trailing-whitespace-robust so a lone " " wrap token keeps its advance on
      // trimming engines (see the `push` measure above): the width fed to the
      // wrap-fit decision must match the width the token is drawn at.
      const tokW = measureAdvanceWithTrailingSpace(ctx, token);
      const isWhitespace = /^\s+$/.test(token);

      // If already in tab mode, collect all text into tabStop.segments (no wrap)
      if (tabActive) {
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        continue;
      }

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
          if (lineW + chW > lineMaxW() && lineW > 0) newLine();
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
      if (hasCJK) {
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
          if (lineW > 0 && lineW + tokenW > lineMaxW()) newLine();
          for (const m of measured) {
            push(m.ch, m.font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, { ...segExtras, fontFamily: m.family });
          }
          continue;
        }
        let rest = measured;
        while (rest.length > 0) {
          const n = fitCjkLine(rest, lineW, lineMaxW(), DEFAULT_KINSOKU_RULES);
          if (n === 0) {
            newLine(); // non-empty line can't take the run head → break, retry empty
            continue;
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

      if (lineW + tokW <= lineMaxW()) {
        push(token, font, sizePx, color, segUnderline, run.strikethrough, run.baseline ?? undefined, segExtras);
        if (isWhitespace) hasWhitespaceOnLine = true;
      } else if (isWhitespace) {
        if (lineW > 0) newLine();
      } else if (tokW > lineMaxW()) {
        if (lineW > 0) newLine();
        for (const ch of token) {
          ctx.font = font;
          const chW = ctx.measureText(ch).width;
          if (lineW + chW > lineMaxW() && lineW > 0) newLine();
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
        // UAX#14 LB13: if the overflowing token is a non-starter (comma, …) glued
        // to the word ending this line (no whitespace between — e.g. authored in a
        // separate run), move the WHOLE word down with it so the comma never leads
        // a line and the word is never torn. Otherwise wrap normally.
        const firstCp = token.codePointAt(0);
        const glued =
          firstCp !== undefined &&
          DEFAULT_KINSOKU_RULES.lineStartForbidden.has(firstCp) &&
          /\S$/.test(currentLine.segments.at(-1)?.text ?? '');
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
      const bitmap = await getCachedBitmapByPath(
        fill.imagePath,
        fill.mimeType,
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

function renderShape(ctx: CanvasRenderingContext2D, el: ShapeElement, scale: number, themeDefaultColor = '#000000', slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, onTextRun?: TextRunCallback, fetchImage?: FetchImage) {
  const x = emuToPx(el.x, scale);
  const y = emuToPx(el.y, scale);
  const w = emuToPx(el.width, scale);
  const h = emuToPx(el.height, scale);

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
      const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
      renderTextBody(ctx, el.textBody, x, y, w, h, scale, defaultTextColor, el.rotation, el.flipH, el.flipV, themeDefaultColor, slideNumber, rc, onTextRun, false, fetchImage);
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
    const defaultTextColor = el.defaultTextColor ? hexToRgba(el.defaultTextColor) : null;
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
    renderTextBody(ctx, el.textBody, tx, ty, tw, th, scale, defaultTextColor, el.rotation, false, false, themeDefaultColor, slideNumber, rc, onTextRun, false, fetchImage);
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
    // After rotation the "width" direction of the new frame is the original height
    renderTextBody(ctx, { ...body, vert: 'horz' }, -bh / 2, -bw / 2, bh, bw, scale, shapeDefaultTextColor, 0, false, false, themeDefaultColor, slideNumber, rc, wrappedOnTextRun, false, fetchImage);
    ctx.restore();
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
      const bSizePx = b.sizePct != null
        ? bulletBaseSizePx * (b.sizePct / 100)
        : bulletBaseSizePx;
      // If the char was mapped to a Unicode symbol, use sans-serif for reliable rendering.
      // Otherwise use the specified font (e.g. Wingdings on systems that have it).
      const convertedFamily = bulletLabel !== b.char ? 'sans-serif' : normalizeFontFamily(b.fontFamily ?? null, rc);
      bulletFont  = buildFont(false, false, bSizePx, convertedFamily, rc);
      bulletColor = b.color ? hexToRgba(b.color) : bulletInheritedColor;
    } else if (bullet.type === 'autoNum') {
      bulletFont  = buildFont(false, false, bulletBaseSizePx, 'sans-serif', rc);
      bulletColor = bulletInheritedColor;
    } else if (bullet.type === 'blip') {
      // ECMA-376 §21.1.2.4.2 picture bullet. The bitmap is drawn as a square
      // sized to the text (the bullet's em box), scaled by `<a:buSzPct>`
      // (§21.1.2.4.3; default 100%). It's not a glyph, so there is no label —
      // an empty paragraph still draws no marker (bulletLabel stays '' and the
      // draw site gates the image on the first line having content).
      const b = bullet;
      const sizePx = b.sizePct != null
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

    // Measure line for alignment AND baseline ascent in one pass.
    // actualBoundingBoxAscent gives the real font ascent for the rendered glyphs,
    // replacing the 0.8×lineHeight heuristic that over-estimates for CJK and
    // tall fonts, causing text to sit too low within the line box.
    let lineWidth = 0;
    let maxAscent = lineHeight * 0.8; // fallback when no segments
    for (const seg of line.segments) {
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

    // Draw bullet. Under RTL the bullet hangs in the right gutter: mirror its
    // LTR offset (textX − bulletX, i.e. the bullet→text gap) about the text
    // column's right edge.
    if (bulletLabel) {
      ctx.font = bulletFont;
      ctx.fillStyle = bulletColor;
      if (paraNeedsBidi && baseRtl) {
        const prevDir = ctx.direction;
        ctx.direction = 'rtl';
        const bulletW = ctx.measureText(bulletLabel).width;
        ctx.fillText(bulletLabel, textX + textMaxW + (textX - bulletX) - bulletW, baseline);
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
          // Mirror into the right gutter, matching the char-bullet RTL offset.
          const imgX = textX + textMaxW + (textX - bulletX) - w;
          ctx.drawImage(bmp, imgX, imgY, w, h);
        } else {
          ctx.drawImage(bmp, bulletX, imgY, w, h);
        }
      }
    }

    const effectiveTextX = textX + textXOffset;
    let penX: number;
    if (alignment === 'ctr') {
      penX = effectiveTextX + (textMaxW - textXOffset - lineWidth) / 2;
    } else if (alignment === 'r') {
      penX = textX + textMaxW - lineWidth;
    } else {
      penX = effectiveTextX;
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
      : alignment === 'dist' || alignment === 'thaiDist' ? 'dist' as const
      : null;
    // A line broken at a right/centre tab stop keeps its pre-tab text natural:
    // justifying it would spread that text across the whole column and overlap
    // the tab-aligned remainder drawn after this loop. Skip justify for it.
    const hasTabStop = !!line.tabStop && line.tabStop.segments.length > 0;
    // A `just` line ended by a manual <a:br> is left-aligned like the last line
    // (§20.1.10.59 + §21.1.2.2.1); `dist` ignores this and fills every line
    // (justifyLine only suppresses the last line for `just`).
    const endsLogicalLine = isLastLine || (line.endsWithBreak ?? false);
    const drawSegs = justifyMode && !paraNeedsBidi && !hasTabStop
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
    if (paraNeedsBidi) ctx.direction = 'ltr'; // reset before tab-stop / next line

    // ── Tab-stop segments (right-aligned or centred at tab stop position) ──
    if (line.tabStop && line.tabStop.segments.length > 0) {
      const tabAbsX = bx + lPad + line.tabStop.px;
      let totalTabW = 0;
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        const ls = seg.letterSpacingPx ?? 0;
        totalTabW += ctx.measureText(seg.text).width + ls * codePointCount(seg.text);
      }
      let tabPenX: number;
      if (line.tabStop.algn === 'r') {
        tabPenX = tabAbsX - totalTabW;
      } else if (line.tabStop.algn === 'ctr') {
        tabPenX = tabAbsX - totalTabW / 2;
      } else {
        tabPenX = tabAbsX;
      }
      for (const seg of line.tabStop.segments) {
        ctx.font = seg.font;
        ctx.fillStyle = seg.color;
        const tabLs = seg.letterSpacingPx ?? 0;
        // Highlight box behind tab-stop-aligned glyphs (ECMA-376 §21.1.2.3.4).
        // No justification stretch on tab-stop runs, so the advance is just
        // glyph measure + letter spacing.
        if (seg.highlight && seg.text) {
          const hlW = ctx.measureText(seg.text).width + tabLs * codePointCount(seg.text);
          paintHighlight(ctx, tabPenX, baseline, hlW, seg.sizePx, seg.highlight, seg.color);
        }
        if (tabLs > 0 && seg.text.length > 1) {
          // rPr @spc (§21.1.2.3.x): draw the whole CONTEXTUALLY-shaped string in
          // ONE fillText with canvas letterSpacing, mirroring drawWithFont (PR
          // #627) / docGrid (docx #626). The old per-glyph `measure(ch)+tabLs`
          // summed ISOLATED advances and overran the contextual box `tabSegW` at
          // 約物半角 punctuation (opening brackets collapse to half-width only in
          // context), pulling later glyphs into the next tab segment. ctx.
          // letterSpacing keeps the drawn advance == measure(seg.text)[contextual]
          // + tabLs·codePointCount == tabSegW.
          const lctx = ctx as CanvasRenderingContext2D & { letterSpacing: string };
          const prev = lctx.letterSpacing;
          try { lctx.letterSpacing = `${tabLs}px`; } catch { /* older engines */ }
          ctx.fillText(seg.text, tabPenX, baseline);
          try { lctx.letterSpacing = prev; } catch { /* ignore */ }
        } else {
          ctx.fillText(seg.text, tabPenX, baseline);
        }
        ctx.font = seg.font;
        const tabSegW = ctx.measureText(seg.text).width + tabLs * codePointCount(seg.text);
        if (onTextRun && seg.text) {
          onTextRun({
            text: seg.text,
            inShapeX: tabPenX - bx,
            inShapeY: cursorY - by,
            w: tabSegW,
            h: lineHeight,
            fontSize: seg.sizePx,
            font: seg.font,
            shapeX: bx,
            shapeY: by,
            shapeW: bw,
            shapeH: bh,
            rotation: shapeRotation,
          });
        }
        tabPenX += tabSegW;
      }
    }

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

function getPosterBitmap(
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
        bitmap = dataIsSvg
          ? await getCachedSvgImageByPath(el.imagePath, fetchImage)
          : await getCachedBitmapByPath(el.imagePath, el.mimeType, fetchImage, { widthPt, heightPt });
      }
    } else if (dataIsSvg) {
      // SVG-only picture (here either because it has a crop, or — defensively —
      // because no svgImagePath was surfaced): decode through the SVG path since
      // createImageBitmap can't.
      bitmap = await getCachedSvgImageByPath(el.imagePath, fetchImage);
    } else {
      bitmap = await getCachedBitmapByPath(el.imagePath, el.mimeType, fetchImage, { widthPt, heightPt });
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

  let drewPoster = false;
  if (el.posterPath && fetchMedia) {
    try {
      // Poster is cached (and prefetched by renderSlide); do not close it here —
      // it is reused across renders of the same slide.
      const bitmap = await getPosterBitmap(el, fetchMedia);
      ctx.drawImage(bitmap, x, y, w, h);
      drewPoster = true;
    } catch {
      // fall through to plain fill
    }
  }
  if (!drewPoster) {
    ctx.fillStyle = el.mediaKind === 'video' ? '#111' : '#f0f0f0';
    ctx.fillRect(x, y, w, h);
  }

  if (skipControls) return;

  drawPlayBadge(ctx, x + w / 2, y + h / 2, w, h, 'paused');
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

function renderTable(ctx: CanvasRenderingContext2D, el: TableElement, scale: number, slideNumber?: number, rc: RenderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 }) {
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
  const jobs: Array<{ cell: TableCell; colX: number; rowY: number; cellW: number; cellH: number }> = [];
  for (let ri = 0; ri < el.rows.length; ri++) {
    const row = el.rows[ri];
    const rowY = rowTop[ri];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      // Merged-cell continuations are drawn by their anchor cell.
      if (cell.hMerge || cell.vMerge) continue;
      const cellW = spannedWidth(ci, cell.gridSpan || 1);
      let cellH = 0;
      for (let span = 0; span < (cell.rowSpan || 1); span++) cellH += rowHeights[ri + span] ?? 0;
      const colX = spannedLeft(ci, cell.gridSpan || 1);
      jobs.push({ cell, colX, rowY, cellW, cellH });
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
  for (const { cell, colX, rowY, cellW, cellH } of jobs) {
    ctx.save();
    if (cell.borderT) {
      applyStroke(ctx, cell.borderT, scale);
      const dy = crispOffset(rowY, ctx.lineWidth, dpr);
      ctx.beginPath();
      ctx.moveTo(colX, rowY + dy);
      ctx.lineTo(colX + cellW, rowY + dy);
      ctx.stroke();
    }
    if (cell.borderB) {
      applyStroke(ctx, cell.borderB, scale);
      const dy = crispOffset(rowY + cellH, ctx.lineWidth, dpr);
      ctx.beginPath();
      ctx.moveTo(colX, rowY + cellH + dy);
      ctx.lineTo(colX + cellW, rowY + cellH + dy);
      ctx.stroke();
    }
    if (cell.borderL) {
      applyStroke(ctx, cell.borderL, scale);
      const dx = crispOffset(colX, ctx.lineWidth, dpr);
      ctx.beginPath();
      ctx.moveTo(colX + dx, rowY);
      ctx.lineTo(colX + dx, rowY + cellH);
      ctx.stroke();
    }
    if (cell.borderR) {
      applyStroke(ctx, cell.borderR, scale);
      const dx = crispOffset(colX + cellW, ctx.lineWidth, dpr);
      ctx.beginPath();
      ctx.moveTo(colX + cellW + dx, rowY);
      ctx.lineTo(colX + cellW + dx, rowY + cellH);
      ctx.stroke();
    }
    // Diagonal borders: top-left→bottom-right and bottom-left→top-right
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
  canvas.width  = canvasW * dpr;
  canvas.height = canvasH * dpr;
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
  ctx.scale(dpr, dpr);

  const rc: RenderContext = {
    themeMajorFont: opts.majorFont ?? null,
    themeMinorFont: opts.minorFont ?? null,
    themeHlinkColor: opts.hlinkColor ?? null,
    dpr,
  };

  await renderBackground(ctx, slide.background, canvasW, canvasH, scale, opts.fetchImage);
  if (superseded()) return canvas;

  // Pre-rasterize any equations so the synchronous text layout can place them.
  // `math` is the engine injected once at PptxPresentation.load and threaded in
  // here; without it, equations are skipped and the asset never enters the bundle.
  if (opts.math) await prepareSlideMath(slide, opts.math);
  if (superseded()) return canvas;

  const themeDefaultColor = opts.defaultTextColor
    ? `#${opts.defaultTextColor}`
    : '#000000';

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
        void getCachedBitmapByPath(p.imagePath, p.mimeType, opts.fetchImage, {
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
    }
  }

  // Hidden-slide dimming (a render mechanism; the caller decides WHEN — see
  // PptxViewer's 'dim' mode). Overlay AFTER all content so it reads faintly.
  if (superseded()) return canvas;
  if (opts.dim) applyDimOverlay(ctx, opts.dim, canvasW, canvasH);

  return canvas;
}
