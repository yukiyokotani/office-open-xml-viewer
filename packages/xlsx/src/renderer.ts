import type {
  Worksheet, Styles, Cell, CellValue, CellFont, CellFill, Border, BorderEdge, CellXf,
  ViewportRange, RenderViewportOptions, XlsxTextRunInfo,
  CfRule, CellRange, CfStop, CfValue, Dxf, Hyperlink, DefinedName,
  Run, GradientFillSpec, ShapeInfo, SlicerItem,
} from './types.js';
import { crispOffset, renderChart, renderSparkline, renderPresetShape, createAuxCanvas, PT_TO_PX, EMU_PER_PX, mathToMathML, recolorSvg, classifyCjkFont, classifyFontGeneric, cjkFallbackChain, NON_CJK_SANS_FALLBACKS, NON_CJK_SERIF_FALLBACKS, kinsokuAdjustedSplit, DEFAULT_KINSOKU_RULES, isCjkBreakChar, isLatinWordCodePoint, xlsxBorderDashArray, drawImageCropped, hexToRgba, intendedSingleLinePx, measureAdvanceWithTrailingSpace, type SparklineModel, type MathNode, type MathRenderer } from '@silurus/ooxml-core';
import { evalFormulaToBool, todaySerial, nowSerial } from './formula.js';
import { formatCellValue } from './number-format.js';
import { type CfContext, compileCf, evaluateCf } from './conditional-format.js';
import { computeLineVisualOrder, cellBaseRtl, resolveCellBidi } from './bidi-line.js';
import { parseA1 } from './a1.js';

// Default font stack. Calibri is the workbook default font in Excel; on
// systems without Office (macOS / Linux) the browser would otherwise fall
// back to Arial / Helvetica, which is meaningfully wider than Calibri at
// every weight/size combination. Carlito is the Google-released, metric-
// compatible Calibri clone (same advance widths and ascender / descender
// metrics) and is loaded opt-in by `XlsxWorkbook.load({ useGoogleFonts:
// true })`. Listing it in the cascade means: Calibri (Windows / Office)
// → Carlito (loaded webfont) → Arial → sans-serif. Caladea is the same
// for Cambria.
// The two trailing Noto Arabic faces are generic Arabic-script fallbacks:
// when the primary Latin faces (Calibri / Carlito / Arial) lack a requested
// glyph, the browser advances down the cascade per-glyph, so any Arabic
// codepoint resolves to a real web font (loaded by `XlsxWorkbook.load`'s
// useGoogleFonts path) instead of an oversized OS Arabic face. Latin glyphs
// still bind to the earlier faces, so Latin rendering is unchanged.
// The trailing non-CJK Noto faces (Hebrew / Thai / Devanagari, plus "Noto Sans"
// for Cyrillic) extend the same per-glyph fallback idea to the other
// non-Latin, non-CJK scripts: any such codepoint resolves to a real web font
// (loaded opt-in via useGoogleFonts) instead of an OS face or tofu. CJK is NOT
// appended here — shared Han glyphs differ in shape per language, so the
// correct Noto CJK is chosen per cell from the cell's font name; see
// fontStackFor() / cssTailFor().
const NON_CJK_SANS_TAIL = NON_CJK_SANS_FALLBACKS.map((n) => `"${n}"`).join(', ');
const NON_CJK_SERIF_TAIL = NON_CJK_SERIF_FALLBACKS.map((n) => `"${n}"`).join(', ');
const DEFAULT_FONT_FAMILY =
  `"Calibri", "Carlito", "Cambria", "Caladea", Arial, "Noto Naskh Arabic", "Noto Sans Arabic", ${NON_CJK_SANS_TAIL}, sans-serif`;
// Serif counterpart of DEFAULT_FONT_FAMILY. A Latin *serif* cell font the host
// lacks (Century, Garamond, …) must degrade to a serif — Excel renders such a
// cell with a serif, not the sans default. Cambria is Office's serif; Caladea is
// its metric-compatible clone (loaded opt-in via useGoogleFonts), then web-safe
// serifs, ending in the `serif` generic.
const DEFAULT_SERIF_FONT_FAMILY =
  `"Cambria", "Caladea", "Times New Roman", "Liberation Serif", "Noto Naskh Arabic", "Noto Sans Arabic", ${NON_CJK_SERIF_TAIL}, serif`;
// Monospace counterpart: a monospaced cell font the host lacks degrades to a
// monospace generic rather than the proportional sans default.
const DEFAULT_MONO_FONT_FAMILY = `"Courier New", "Liberation Mono", monospace`;

/**
 * CSS font-family TAIL (everything after the cell's named face) for an xlsx
 * cell. For a CJK cell font the matching Noto CJK leads (so shared Han glyphs
 * take the document language's shapes; see core/fonts/scripts.ts), followed by
 * the standard Latin/Arabic/non-CJK fallbacks. A non-CJK cell font picks the
 * default chain by its generic class ({@link classifyFontGeneric}) so a Latin
 * serif/mono face the host lacks degrades to the matching generic. Exported for
 * unit testing.
 */
export function cssTailFor(name: string | null | undefined): string {
  const cjk = name ? classifyCjkFont(name) : null;
  const generic = classifyFontGeneric(name); // 'serif' | 'sans' | 'mono'
  if (!cjk) {
    // Non-CJK (Latin) cell font: choose the default chain by generic class so a
    // Latin serif/mono face the host lacks degrades to the matching generic
    // (Excel renders serif/mono, not the sans default).
    if (generic === 'serif') return DEFAULT_SERIF_FONT_FAMILY;
    if (generic === 'mono') return DEFAULT_MONO_FONT_FAMILY;
    return DEFAULT_FONT_FAMILY;
  }
  const serif = generic === 'serif';
  const cjkPart = cjkFallbackChain(cjk, serif ? 'serif' : 'sans')
    .map((n) => `"${n}"`)
    .join(', ');
  const tail = serif ? NON_CJK_SERIF_TAIL : NON_CJK_SANS_TAIL;
  const genericKeyword = serif ? 'serif' : 'sans-serif';
  // CJK Noto leads, then Latin/metric substitutes, Arabic, non-CJK scripts.
  return `${cjkPart}, "Calibri", "Carlito", "Cambria", "Caladea", Arial, "Noto Naskh Arabic", "Noto Sans Arabic", ${tail}, ${genericKeyword}`;
}

/** Full CSS font-family list for a cell font name (named face first). */
export function fontStackFor(name: string | null | undefined): string {
  return name ? `"${name}", ${cssTailFor(name)}` : DEFAULT_FONT_FAMILY;
}

const DEFAULT_FONT_SIZE = 11;
// Fallback Max Digit Width of the Normal-style font when the workbook's
// default font isn't known. Calibri 11 pt at 96 DPI ≈ 8 px (Canvas2D
// measurement), matching the EMU offsets Excel 365 writes into
// <xdr:twoCellAnchor>. ECMA-376 §18.3.1.13 defines MDW as the maximum
// rendered width among the digits 0-9 in the workbook's Normal-style font,
// so the spec-correct value depends on which font and point size that style
// resolves to (e.g. Meiryo UI 10 pt yields MDW ≈ 6 px).
const MDW_FALLBACK = 8;

export const HEADER_W = 50;
export const HEADER_H = 22;

/**
 * Sheet-level right-to-left horizontal mirror (ECMA-376 §18.3.1.87
 * `<sheetView rightToLeft>`). The grid is always laid out left-to-right
 * internally; an RTL sheet is produced by mirroring every horizontal extent
 * about the canvas width. A left-anchored rect `[x, x + w]` maps to
 * `[canvasW - x - w, canvasW - x]`.
 *
 * Because the LTR layout puts the row-header strip at `[0, HEADER_W]` and the
 * cell area at `[HEADER_W, canvasW]`, mirroring moves the header strip to the
 * right edge and the cell area to `[0, canvasW - HEADER_W]`, matching Excel.
 *
 * This is the SINGLE source of truth for the RTL x transform. The Canvas
 * renderer, the selection overlay, and pointer hit-testing must all use it so
 * that a cell drawn at screen-x is the same cell a click at screen-x resolves
 * to, at every scroll offset. `canvasW` is the CSS-pixel width of the drawing
 * surface (`canvasArea.clientWidth`), identical to `ctx.canvas.width / dpr`.
 *
 * The transform is an involution: applying it to a screen point recovers the
 * logical-LTR point (`rtlMirrorX(rtlMirrorX(x, w, W), w, W) === x`), so the
 * same function serves both cell→px (draw) and px→cell (hit-test, with w = 0
 * for a point).
 */
export function rtlMirrorX(x: number, w: number, canvasW: number): number {
  return canvasW - x - w;
}

// Thin line drawn between frozen and scrollable areas
const FREEZE_LINE_COLOR = '#7a7a7a';

/** Cache of Max Digit Width per "family:sizePt" key. The Canvas2D
 *  `measureText` call is cheap but not free; column-width conversion is
 *  invoked many times per render so we memoize. */
const mdwCache = new Map<string, number>();

/** Excel's empirical MDW overrides for fonts that the host might not have
 *  installed (e.g. Meiryo UI on macOS, where Canvas2D falls back to a
 *  narrower sans-serif and undermeasures the digits — verified against
 *  private/sample-10's column widths in Excel: 21.125 chars = 169 px,
 *  which requires MDW=8). Without these the rendered column widths drift
 *  smaller than Excel's, which then offsets every drawing anchor inside
 *  the sheet (sample-10's H7 sun ended up one cell to the right of where
 *  Excel renders it).
 *
 *  Only listed when the Canvas2D fallback measurement diverges from
 *  Excel's actual MDW; other fonts (e.g. Yu Gothic 12 pt where Canvas
 *  happens to land on 8) continue to use the measurement. */
const MDW_TABLE: Record<string, Record<number, number>> = {
  'meiryo ui':       { 10: 8, 11: 8 },
  'meiryo':          { 10: 8, 11: 8 },
};

/** Measure the Max Digit Width (ECMA-376 §18.3.1.13) for an arbitrary font
 *  using Canvas2D. The maximum of `measureText('0'..'9').width` is taken,
 *  rounded to the nearest pixel to match Excel's storage of integer pixel
 *  widths in `<col>` width values.
 *
 *  When the requested font isn't installed on the host (e.g. Meiryo UI on
 *  macOS), Canvas2D silently falls back to a narrower sans-serif face and
 *  the digit width comes back ~1 px too small. That offset cascades into
 *  the column widths, shifting every drawing anchor inside the sheet
 *  relative to where Excel placed it (private/sample-10 H7 sun ended up
 *  one cell to the right of where Excel renders it). So we consult a small
 *  lookup table of Excel's documented MDW values first and only fall back
 *  to measurement for unknown faces. */
export function computeMdw(family: string, sizePt: number): number {
  const key = `${family}:${sizePt}`;
  const cached = mdwCache.get(key);
  if (cached !== undefined) return cached;
  const tableHit = MDW_TABLE[family.toLowerCase()]?.[Math.round(sizePt)];
  if (tableHit !== undefined) {
    mdwCache.set(key, tableHit);
    return tableHit;
  }
  const sizePx = sizePt * PT_TO_PX;
  // Off-DOM canvas: avoids touching the document tree from background calls.
  const canvas = (typeof OffscreenCanvas !== 'undefined')
    ? new OffscreenCanvas(1, 1)
    : (typeof document !== 'undefined' ? document.createElement('canvas') : null);
  if (!canvas) return MDW_FALLBACK;
  const ctx = canvas.getContext('2d');
  if (!ctx) return MDW_FALLBACK;
  // Quote the family so multi-word names like "Meiryo UI" parse as one face.
  ctx.font = `${sizePx}px ${fontStackFor(family)}`;
  let mdw = 0;
  for (const d of '0123456789') {
    const w = ctx.measureText(d).width;
    if (w > mdw) mdw = w;
  }
  const out = Math.round(mdw) || MDW_FALLBACK;
  mdwCache.set(key, out);
  return out;
}

/** Resolve the Max Digit Width for a worksheet's Normal-style font. Falls
 *  back to the Calibri 11 pt baseline (~8 px) when the parser couldn't
 *  determine the workbook's default font. */
export function getMdwForWorksheet(ws: { defaultFontFamily?: string; defaultFontSize?: number }): number {
  if (!ws.defaultFontFamily || !ws.defaultFontSize) return MDW_FALLBACK;
  return computeMdw(ws.defaultFontFamily, ws.defaultFontSize);
}

/** Convert a stored column-width value (ECMA-376 §18.3.1.13 `<col width>`, in
 *  "number of characters" = max digit widths) to CSS pixels.
 *
 *  This is the spec's file→pixel formula verbatim:
 *    `Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)`
 *  Note both truncations: the `Truncate(128 / MDW)` constant is computed and
 *  truncated *before* it is folded into the numerator (§18.3.1.13), then the
 *  whole expression is truncated to an integer pixel. Excel stores integer
 *  pixel column widths, so this yields exactly the width Excel renders. */
export function colWidthToPx(w: number, mdw: number = MDW_FALLBACK): number {
  return Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** Analytic inverse of {@link colWidthToPx}: the internal column-width value (in
 *  "max digit widths") that renders back to *exactly* `px` logical pixels, so a
 *  column dragged to N px paints at N px with no drift (WYSIWYG). Used only by
 *  the drag-to-resize handles (issue #567) to write the dragged size into the
 *  in-memory worksheet model.
 *
 *  This is deliberately NOT the ECMA-376 §18.3.1.13 file px→character formula
 *  `Truncate((px - 5) / MDW * 100 + 0.5) / 100`, for two reasons:
 *   (a) this viewer never serializes the workbook, so the model's width unit is
 *       purely internal — the only contract it must honor is the exact round-trip
 *       with {@link colWidthToPx}, which `px / MDW` satisfies (the constant
 *       `Truncate(128/MDW)/256 * MDW` added by the forward formula stays in
 *       `[0, 1)`, so `trunc(px + c) === px` for integer `px`); and
 *   (b) the spec formula degenerates below its hard-coded 5 px cell padding —
 *       for small dragged columns `(px - 5)` goes to zero or negative, producing
 *       0 / negative character widths that would make drags snap or collapse.
 *
 *  If a file-export path is ever added, do NOT reuse this for serialization:
 *  switch to the spec px→character formula above AND record `customWidth="1"`
 *  provenance (§18.3.1.13) at the serialization boundary. */
export function pxToColWidth(px: number, mdw: number = MDW_FALLBACK): number {
  return px / mdw;
}

/** Convert a row height value from the parser into CSS pixels.
 *
 * ECMA-376 §18.3.1.73 (`<row ht>`) and §18.3.1.81 (`sheetFormatPr@defaultRowHeight`)
 * both specify the value in points. Convert pt → CSS px at 96 DPI (×4/3)
 * to match what Excel actually displays. The parser keeps both per-row
 * heights and the intrinsic default in points so this single conversion
 * applies to either source. */
export function rowHeightToPx(h: number): number {
  return Math.round(h * PT_TO_PX);
}

/** Inverse of {@link rowHeightToPx}: the row height in points that renders back
 *  to exactly `px` logical pixels (`round(px/PT_TO_PX * PT_TO_PX) === px`). Used
 *  by the drag-to-resize handles (issue #567). */
export function pxToRowHeight(px: number): number {
  return px / PT_TO_PX;
}

/**
 * Fill a data-bar rectangle. Excel 2010+ dataBars default to a horizontal
 * gradient (`x14:dataBar@gradient="1"`): solid color on the left, fading
 * to an ~85%-tinted-to-white version on the right. We render with alpha
 * stops rather than literally mixing toward white so underlying cell
 * background (including zebra-striping or fills) shows through. With
 * `gradient="0"` the bar is drawn as a flat solid color.
 */
function fillDataBar(
  ctx: CanvasRenderingContext2D,
  color: string,
  x: number, y: number, w: number, h: number,
  gradient: boolean,
): void {
  if (w <= 0 || h <= 0) return;
  if (gradient) {
    const grad = ctx.createLinearGradient(x, y, x + w, y);
    grad.addColorStop(0, hexToRgba(color, 0.85));
    grad.addColorStop(1, hexToRgba(color, 0.15));
    ctx.fillStyle = grad;
  } else {
    ctx.fillStyle = hexToRgba(color);
  }
  ctx.fillRect(x, y, w, h);
}

/**
 * Fractional fg coverage for an ECMA-376 ST_PatternType (§18.8.22). Values are
 * derived from the spec's verbal descriptions — gray125 = 12.5% fg on bg,
 * gray0625 = 6.25%, mediumGray ≈ 50%, darkGray ≈ 75%, lightGray ≈ 25%. Hatch
 * variants (darkHorizontal etc.) approximate their average ink density.
 * Unknown values default to 1 so they render as solid fg.
 */
function patternCoverage(pt: string): number {
  switch (pt) {
    case 'solid':         return 1;
    case 'darkGray':      return 0.75;
    case 'mediumGray':    return 0.50;
    case 'lightGray':     return 0.25;
    case 'gray125':       return 0.125;
    case 'gray0625':      return 0.0625;
    // Directional hatches: visual ink ratio is roughly 50/50 at default cell
    // sizes; without a true hatch tile we blend at 50% so the cell reads as a
    // middle-tone fill rather than a solid fg block.
    case 'darkHorizontal':
    case 'darkVertical':
    case 'darkDown':
    case 'darkUp':
    case 'darkGrid':
    case 'darkTrellis':   return 0.5;
    case 'lightHorizontal':
    case 'lightVertical':
    case 'lightDown':
    case 'lightUp':
    case 'lightGrid':
    case 'lightTrellis':  return 0.25;
    default:              return 1;
  }
}

/**
 * Cache of (pattern, fg, bg, ctx scale) → CanvasPattern. Keyed by a compound
 * string so the same pattern across thousands of cells only builds the tile
 * once *for a given destination scale*. The ctx scale is part of the key
 * because pat.setTransform() pre-bakes the inverse-scale transform into the
 * pattern, and we don't want a Storybook scale=1.5 cell to reuse a cached
 * pattern from a previous scale=2 render.
 */
const PATTERN_CACHE = new Map<string, CanvasPattern | null>();

/**
 * Bitmaps for ECMA-376 ST_PatternType (§18.8.22). Each entry is a square tile
 * — the row count (rows.length) is the tile size in pixels, and bit
 * (size - 1) of each row is the leftmost pixel. A `1` bit places fg on the
 * tile, `0` leaves bg.
 *
 * Mixing tile sizes lets us pick the smallest square that exactly fits each
 * pattern's natural pitch: gray and diagonal families fit in 8×8, but the
 * horizontal / vertical / grid families need a 12×12 tile so dark and light
 * variants can share the same line pitch (3 px) while differing only in
 * stroke thickness (2 px vs 1 px). At 8×8 this would have forced a 4-px
 * pitch on dark vs 2-px on light — visibly different line counts instead of
 * the matched-count, different-thickness pair Excel renders.
 *
 * Values follow the long-standing Office pattern set; the spec only names
 * the patterns, never their pixel geometry.
 *
 * Coverage targets:
 *   gray125    12.5%   sparse dot  (1 dot per 8 pixels)
 *   gray0625    6.25%   very sparse (1 dot per 16 pixels)
 *   lightGray  25%      checker dot (1 in 4)
 *   mediumGray 50%      full checker
 *   darkGray   75%      inverse of lightGray
 */
const PATTERN_BITMAPS: Record<string, number[]> = {
  // ── 8×8 — gray-family dot patterns ────────────────────────────────────
  // Constructed as a single dot motif scaled up tier-by-tier — each pattern
  // is the previous one with extra dots interleaved at the same 4-row pitch.
  // gray0625 is the reference seed (4 dots / 64 ≈ 6%) and the cascade roughly
  // doubles density at each step, so all five tiers read as the same stipple
  // texture at progressively higher coverage. Some moiré at intermediate
  // zooms is acceptable per user preference; the visual continuity wins.
  gray0625:   [0b10000000, 0b00000000, 0b00001000, 0b00000000, 0b10000000, 0b00000000, 0b00001000, 0b00000000], // ≈ 6%
  gray125:    [0b10001000, 0b00000000, 0b00100010, 0b00000000, 0b10001000, 0b00000000, 0b00100010, 0b00000000], // ≈ 12%
  lightGray:  [0b10101010, 0b00000000, 0b01010101, 0b00000000, 0b10101010, 0b00000000, 0b01010101, 0b00000000], // ≈ 25%
  mediumGray: [0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101], // ≈ 50%
  // 75% — distribute the empty pixels evenly across rows so the cell reads
  // as a solid darker grey instead of alternating bands.
  darkGray:   [0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101, 0b01110111, 0b11011101], // ≈ 75%

  // ── 12×12 — horizontal / vertical (matched line count) ───────────────
  // dark* and light* share the same 4-line-per-tile count and 3-px pitch;
  // they differ only in stroke thickness — dark uses a 2-px bar and a
  // 1-px gap, light uses a 1-px line and a 2-px gap. Per user feedback
  // the visual rule is "B19 has the same number of horizontal lines as
  // B18, just thinner", so the matched-count construction is restored
  // here. The earlier "more lines in light" 8×8 attempt was Excel-wrong.
  darkHorizontal: [
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
    0b111111111111, 0b111111111111, 0b000000000000,
  ],
  lightHorizontal: [
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
    0b111111111111, 0b000000000000, 0b000000000000,
  ],
  // darkVertical: 2-col bars at cols 0,1 / 3,4 / 6,7 / 9,10
  //   bits set per row: 11,10 + 8,7 + 5,4 + 2,1 = 0b110110110110 = 0xDB6
  darkVertical:   Array(12).fill(0xDB6),
  // lightVertical: 1-col lines at cols 0 / 3 / 6 / 9
  //   bits set per row: 11 + 8 + 5 + 2 = 0b100100100100 = 0x924
  lightVertical:  Array(12).fill(0x924),

  // darkGrid: Excel renders this as a fine 2×2 checkerboard, not a thick
  // horizontal+vertical lattice. Each black/white "cell" is 2×2 source
  // pixels, so the cell reads as a clear "市松模様" rather than a solid
  // 50% gray (the 1×1 checkerboard mediumGray we already ship).
  darkGrid:  [0b11001100, 0b11001100, 0b00110011, 0b00110011, 0b11001100, 0b11001100, 0b00110011, 0b00110011],
  // lightGrid: a sparse grid with 4-px pitch — horizontal lines at rows
  // 0 and 4, vertical lines at cols 0 and 4. Grid cells (white squares
  // between the lines) are 3×3 source pixels, matching Excel's rendering
  // of lightGrid where the lattice is clearly larger than the dense
  // version and a true "格子" pattern is recognisable.
  lightGrid: [0b11111111, 0b10001000, 0b10001000, 0b10001000, 0b11111111, 0b10001000, 0b10001000, 0b10001000],

  // ── 8×8 — diagonals & trellis ─────────────────────────────────────────
  // darkDown / darkUp: 2-px-wide diagonal stripes every 4 cells.
  // lightDown / lightUp: 1-px-wide diagonal stripes at the same 4-cell pitch.
  darkDown:   [0b11001100, 0b01100110, 0b00110011, 0b10011001, 0b11001100, 0b01100110, 0b00110011, 0b10011001],
  lightDown:  [0b10001000, 0b01000100, 0b00100010, 0b00010001, 0b10001000, 0b01000100, 0b00100010, 0b00010001],
  darkUp:     [0b00110011, 0b01100110, 0b11001100, 0b10011001, 0b00110011, 0b01100110, 0b11001100, 0b10011001],
  lightUp:    [0b00010001, 0b00100010, 0b01000100, 0b10001000, 0b00010001, 0b00100010, 0b01000100, 0b10001000],
  // darkTrellis = darkDown | darkUp.
  darkTrellis:  [0b11111111, 0b01100110, 0b11111111, 0b10011001, 0b11111111, 0b01100110, 0b11111111, 0b10011001],
  // lightTrellis = lightDown | lightUp.
  lightTrellis: [0b10011001, 0b01100110, 0b01100110, 0b10011001, 0b10011001, 0b01100110, 0b01100110, 0b10011001],
};

/**
 * Build a repeating 8x8 tile for an ECMA-376 preset pattern. Returns null
 * for unknown pattern names so the caller can fall back to a flat colour.
 */
function hatchPattern(
  ctx: CanvasRenderingContext2D,
  pt: string,
  fgHex: string,
  bgHex: string,
): CanvasPattern | null {
  // Read the destination context's effective scale so the offscreen tile is
  // sized at the destination's device-pixel resolution. Drawing the source
  // pre-scaled keeps each "1 bit" at exactly one destination *device* pixel
  // when the rendering context is doing ctx.scale(dpr * cs) for hi-DPI /
  // Storybook zoom — without this the crisp 1-px source bits get linearly
  // resampled to 1.5–2 device px and either chunky (integer scale) or
  // blurred-out-of-existence (non-integer scale, the prior pat.setTransform
  // approach lost lightHorizontal entirely at scale=1.5).
  const t = ctx.getTransform();
  const sx = Math.max(1, Math.round(Math.hypot(t.a, t.b)));
  const sy = Math.max(1, Math.round(Math.hypot(t.c, t.d)));
  const key = `${pt}|${fgHex}|${bgHex}|${sx}|${sy}`;
  if (PATTERN_CACHE.has(key)) return PATTERN_CACHE.get(key)!;

  const rows = PATTERN_BITMAPS[pt];
  if (!rows) {
    PATTERN_CACHE.set(key, null);
    return null;
  }

  // Square tile — size inferred from the row count. Bit (size-1) is
  // leftmost so the binary literals read left-to-right at any tile width.
  // The offscreen canvas is sized at sx × tileSize device pixels so each
  // "1 bit" becomes a sx × sy rect when the destination context's scale
  // multiplies it back. Result: the source bit lands on exactly one
  // destination device pixel regardless of the user's CSS zoom.
  const tile = rows.length;
  const off = createAuxCanvas(tile, tile);
  if (!off) { PATTERN_CACHE.set(key, null); return null; }
  const octx = off.getContext('2d') as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (!octx) { PATTERN_CACHE.set(key, null); return null; }

  octx.fillStyle = hexToRgba(bgHex);
  octx.fillRect(0, 0, tile, tile);
  octx.fillStyle = hexToRgba(fgHex);
  for (let y = 0; y < tile; y++) {
    const row = rows[y];
    for (let x = 0; x < tile; x++) {
      if (row & (1 << (tile - 1 - x))) octx.fillRect(x, y, 1, 1);
    }
  }

  const pat = ctx.createPattern(off, 'repeat');
  // Pre-bake an inverse-scale matrix only when the scale rounds to an
  // integer (1, 2, 3, …). Fractional inverses (e.g. 1/1.5 ≈ 0.67) trigger
  // the canvas's bilinear pattern resampler and smear 1-px features into
  // a uniform half-tone — which is exactly what wiped out lightHorizontal
  // at Storybook scale=1.5. For non-integer ctx scales we leave the tile
  // unscaled and accept ~1.5-device-px bits (slightly chunky but crisp).
  if (pat && typeof DOMMatrix !== 'undefined' && (sx >= 2 || sy >= 2)) {
    const m = new DOMMatrix();
    m.scaleSelf(1 / sx, 1 / sy);
    pat.setTransform(m);
  }
  PATTERN_CACHE.set(key, pat);
  return pat;
}

/**
 * Paint a cell background according to its <patternFill> / <gradientFill>.
 *
 * ECMA-376 §18.8.20 / §18.8.22 specifies fg/bg defaults: when the colour
 * children are absent, fg defaults to the system foreground (black) and bg
 * to the system background (white). Without those defaults the directional
 * hatches that Excel emits with no explicit colours (`<patternFill
 * patternType="darkHorizontal"/>` etc.) would render as nothing because
 * the prior gate required a non-null fgColor.
 *
 * Returns true when the cell was painted so the caller can short-circuit
 * its tableStyle / banded fallbacks.
 */
function paintCellPatternFill(
  ctx: CanvasRenderingContext2D,
  fill: CellFill,
  x: number, y: number, w: number, h: number,
): boolean {
  if (fill.gradient && fill.gradient.stops.length > 0) {
    ctx.fillStyle = buildGradientFill(ctx, fill.gradient, x, y, w, h);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const pt = fill.patternType;
  if (!pt || pt === 'none') return false;
  const fg = fill.fgColor ?? '000000';
  const bg = fill.bgColor ?? 'FFFFFF';
  if (pt === 'solid') {
    ctx.fillStyle = hexToRgba(fg);
    ctx.fillRect(x, y, w, h);
    return true;
  }
  const hatch = hatchPattern(ctx, pt, fg, bg);
  if (hatch) {
    ctx.fillStyle = hatch;
  } else {
    const coverage = patternCoverage(pt);
    ctx.fillStyle = coverage >= 1
      ? hexToRgba(fg)
      : blendHex(fg, bg, coverage);
  }
  ctx.fillRect(x, y, w, h);
  return true;
}

/**
 * Build a Canvas gradient object for an xlsx `<gradientFill>`. Linear uses
 * the degree attribute (0° = left→right, 90° = top→bottom). Path gradients
 * radiate from a rectangular inner bounds defined by left/right/top/bottom
 * as fractions of the cell.
 */
function buildGradientFill(
  ctx: CanvasRenderingContext2D,
  g: GradientFillSpec,
  x: number, y: number, w: number, h: number,
): CanvasGradient {
  let grad: CanvasGradient;
  if (g.gradientType === 'path') {
    // Use the inner rectangle's center as the radial origin; radius spans to
    // the farthest cell corner so stop=1 always reaches a cell edge.
    const cxg = x + w * (g.left + (1 - g.right - g.left) / 2);
    const cyg = y + h * (g.top + (1 - g.bottom - g.top) / 2);
    const r = Math.hypot(Math.max(cxg - x, x + w - cxg), Math.max(cyg - y, y + h - cyg));
    grad = ctx.createRadialGradient(cxg, cyg, 0, cxg, cyg, r);
  } else {
    // Linear: rotate around the cell's center and extend to the bounds.
    const rad = (g.degree * Math.PI) / 180;
    const cxg = x + w / 2;
    const cyg = y + h / 2;
    const ext = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
    grad = ctx.createLinearGradient(
      cxg - Math.cos(rad) * ext, cyg - Math.sin(rad) * ext,
      cxg + Math.cos(rad) * ext, cyg + Math.sin(rad) * ext,
    );
  }
  for (const stop of g.stops) {
    const pos = Math.min(1, Math.max(0, stop.position));
    grad.addColorStop(pos, hexToRgba(stop.color));
  }
  return grad;
}

/** Parse an A1-style cell reference to 1-based row/col. Aliased to the shared
 *  {@link parseA1} so the renderer, data-validation and the comment popup all
 *  agree on ref parsing (the shared form also tolerates `$`-absolute markers). */
const parseA1Ref = parseA1;

/**
 * Draw Excel's comment marker — a small filled triangle in the top-right
 * corner of the cell — coloured like Excel's default red indicator. Scales
 * with cell size but is clamped so it stays legible at small zoom.
 */
function drawCommentMarker(
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): void {
  const size = Math.max(4, Math.min(8, Math.min(w, h) * 0.18));
  ctx.save();
  ctx.fillStyle = '#D40000';
  ctx.beginPath();
  ctx.moveTo(x + w - size, y);
  ctx.lineTo(x + w, y);
  ctx.lineTo(x + w, y + size);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** Linear-interpolate two #RRGGBB values in RGB space at coverage fg weight. */
function blendHex(fgHex: string, bgHex: string, fgCoverage: number): string {
  const fh = fgHex.replace('#', '');
  const bh = bgHex.replace('#', '');
  const fr = parseInt(fh.slice(0, 2), 16);
  const fg = parseInt(fh.slice(2, 4), 16);
  const fb = parseInt(fh.slice(4, 6), 16);
  const br = parseInt(bh.slice(0, 2), 16);
  const bg = parseInt(bh.slice(2, 4), 16);
  const bb = parseInt(bh.slice(4, 6), 16);
  const c = Math.min(1, Math.max(0, fgCoverage));
  const r = Math.round(fr * c + br * (1 - c));
  const g = Math.round(fg * c + bg * (1 - c));
  const b = Math.round(fb * c + bb * (1 - c));
  return `rgb(${r},${g},${b})`;
}

/** Vertical pixel metric for cell text: point size → device px at the current
 *  cell scale `cs`. `factor` is the line-height / char-height multiplier (1.2
 *  for wrapped lines, 1.1 for stacked chars, 1.0 for decoration/baseline
 *  offsets). Centralizes the `* cs` factor so a new vertical-metric draw site
 *  can't silently omit it. (Glyph SIZE uses buildFont's floored variant.)
 *
 *  `family` is passed ONLY at the single-line-height sites (factor 1.2) so the
 *  result is floored to the DOCUMENT font's design single-line height (ECMA-376
 *  §17.3.1.33, shared with docx/pptx via core's `intendedSingleLinePx`): Excel
 *  sizes single spacing as a flat 1.2×em, which understates a SUBSTITUTED
 *  Meiryo (1.596×em) / Sakkal Majalla (1.3965×em) line box and makes rows/lines
 *  too short. This is a FLOOR — `intendedSingleLinePx` returns 0 for every
 *  non-tabled family, so `max(base, 0) = base` leaves all other fonts on
 *  Excel's 1.2×em. Do NOT pass `family` at the base-size (no factor) or the 1.1
 *  super/sub sites — those must stay on the flat metric. */
function vMetricPx(sizePt: number, cs: number, factor = 1, family?: string): number {
  const base = Math.round(sizePt * PT_TO_PX * factor * cs);
  if (!family) return base;
  return Math.max(base, Math.round(intendedSingleLinePx(family, sizePt * PT_TO_PX * cs)));
}

function buildFont(font: CellFont, cs = 1): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? 'bold ' : '';
  const sizePx = Math.max(1, Math.round(font.size * PT_TO_PX * cs));
  return `${style}${weight}${sizePx}px ${fontStackFor(font.name)}`;
}

/**
 * Stroke a single or double horizontal text-decoration line
 * (underline / strikethrough). ECMA-376 §18.4.13 ST_UnderlineValues
 * "double" / "doubleAccounting" both render as two parallel lines with a
 * small gap. The two strokes straddle the requested baseline ±1px so the
 * pair reads as a ~3px-thick rule and the visual centre matches Excel.
 */
function drawTextDecoLine(
  ctx: CanvasRenderingContext2D,
  x1: number, x2: number, y: number,
  color: string, double: boolean,
  dpr = 1,
): void {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  // crispOffset snaps each horizontal stroke onto the device pixel grid so the
  // 0.5-px rule renders crisp at any dpr (matches docx/pptx underline/strike).
  if (double) {
    const yUp = y - 1, yDn = y + 1;
    const yu = yUp + crispOffset(yUp, 0.5, dpr);
    const yd = yDn + crispOffset(yDn, 0.5, dpr);
    ctx.moveTo(x1, yu); ctx.lineTo(x2, yu);
    ctx.moveTo(x1, yd); ctx.lineTo(x2, yd);
  } else {
    const ys = y + crispOffset(y, 0.5, dpr);
    ctx.moveTo(x1, ys); ctx.lineTo(x2, ys);
  }
  ctx.stroke();
  ctx.restore();
}

/**
 * Resolve a Run's font against a base Font. Per ECMA-376, a run's <rPr>
 * completely specifies bold/italic/underline/strike for that run, while
 * size/color/name fall back to the base when omitted. A run with no
 * <rPr> (run.font undefined) inherits the base entirely.
 */
function applyRunFont(base: CellFont, run: Run): CellFont {
  const rf = run.font;
  if (!rf) return base;
  return {
    bold: rf.bold,
    italic: rf.italic,
    underline: rf.underline,
    underlineStyle: rf.underlineStyle,
    strike: rf.strike,
    size: rf.size ?? base.size,
    color: rf.color ?? base.color,
    name: rf.name ?? base.name,
    vertAlign: rf.vertAlign,
  };
}

function resolveXf(styles: Styles, styleIndex: number): { font: CellFont; fill: CellFill; border: Border; xf: CellXf } {
  const xf: CellXf = styles.cellXfs[styleIndex] ?? styles.cellXfs[0] ?? {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
  };
  const font: CellFont = styles.fonts[xf.fontId] ?? { bold: false, italic: false, underline: false, strike: false, size: DEFAULT_FONT_SIZE, color: null, name: null };
  const fill: CellFill = styles.fills[xf.fillId] ?? { patternType: 'none', fgColor: null, bgColor: null };
  const border: Border = styles.borders[xf.borderId] ?? { left: null, right: null, top: null, bottom: null };
  return { font, fill, border, xf };
}

function wrapTextLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Hard line breaks (\n from Alt+Enter) always split regardless of wrapText.
  for (const paragraph of text.split('\n')) {
    lines.push(...wrapParagraphLines(ctx, paragraph, maxWidth));
  }
  return lines;
}

/**
 * Apply Japanese line-breaking (kinsoku, 禁則処理) at a wrap boundary.
 *
 * When the wrapper decides to break, it has the code points already committed
 * to the line being closed (`lineCps`) and the code points that will lead the
 * next line (`nextCps`, the overflowing token). Excel — like Word and
 * PowerPoint — forbids a wrapped line from STARTING with a 行頭禁則 char
 * (、。」）…) or ENDING with a 行末禁則 char (「（…), per ECMA-376
 * §17.15.1.58–.60. We delegate the retraction to the shared core engine
 * `kinsokuAdjustedSplit`, which pulls the offending boundary's preceding code
 * point(s) down onto the next line (追い出し).
 *
 * Returns the number of trailing code points of `lineCps` that must move down
 * to lead the next line (ahead of `nextCps`). `0` means the greedy break was
 * already legal — so plain CJK with no forbidden chars at the boundary is
 * unchanged (no regression). `minSplit = 1` keeps ≥1 code point on the closed
 * line; an all-forbidden run falls back to no retraction (never empties a line,
 * never hangs).
 */
function kinsokuRetractCount(lineCps: string[], nextCps: string[]): number {
  if (lineCps.length === 0 || nextCps.length === 0) return 0;
  const combined = [...lineCps, ...nextCps];
  const splitAt = lineCps.length;
  const adj = kinsokuAdjustedSplit(combined, splitAt, DEFAULT_KINSOKU_RULES, 1);
  return splitAt - adj;
}

/**
 * Extend a per-code-point kinsoku retract so it never TEARS a Latin word.
 *
 * `kinsokuRetractCount` retracts glyph-by-glyph — correct for CJK, where every
 * character is a break opportunity, but wrong for Latin: a non-starter (comma,
 * period, …, UAX#14 LB13) overflowing after "system" retracts a single "m" →
 * "syste" / "m,". Latin has no mid-word break opportunity, so when the retract
 * boundary sits between two {@link isLatinWordCodePoint} characters, pull it back
 * to the last whitespace (the real break) so the WHOLE word moves down ahead of
 * the non-starter. If the line is one unbroken word (no whitespace to retract
 * to), keep the original retract rather than empty the line — an over-long
 * single word is handled by the normal overflow path. NOTE: the retract is still
 * capped at the last segment by the caller, so a word split across runs by a
 * formatting change splits at that seam (the comma stays glued to the tail).
 */
function extendLatinWordRetract(lineCps: string[], retract: number): number {
  let r = retract;
  while (r < lineCps.length) {
    const keep = lineCps[lineCps.length - r - 1]; // last char staying on the line
    const move = lineCps[lineCps.length - r];     // first char moving down
    const keepCp = keep?.codePointAt(0);
    const moveCp = move?.codePointAt(0);
    if (keepCp !== undefined && moveCp !== undefined
        && isLatinWordCodePoint(keepCp) && isLatinWordCodePoint(moveCp)) r++;
    else break;
  }
  return r >= lineCps.length ? retract : r;
}

/** Word-wrap a single paragraph (no embedded \n). Unlike a naive
 *  `split(' ')`, CJK characters are treated as individual break opportunities
 *  so that Japanese headings like "夏休みアクティビティ カレンダー 2026"
 *  actually wrap inside a merged cell. ECMA-376 doesn't spec the break
 *  algorithm but this matches what Excel renders on the same input.
 *
 *  At each break we additionally apply kinsoku (`kinsokuRetractCount`) so a
 *  wrapped line never starts with 、。」 or ends with 「（ (ECMA-376
 *  §17.15.1.58–.60), matching Excel's East-Asian wrapping. */
export function wrapParagraphLines(ctx: CanvasRenderingContext2D, paragraph: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Tokenise: runs of non-space non-CJK, single ASCII-space runs, individual
  // CJK characters. Then greedy-fit each token onto the current line.
  const tokens: string[] = [];
  let i = 0;
  while (i < paragraph.length) {
    const ch = paragraph[i];
    const cp = ch.codePointAt(0) ?? 0;
    if (isCjkBreakChar(cp)) {
      tokens.push(ch);
      i += cp > 0xFFFF ? 2 : 1;
    } else if (ch === ' ') {
      let j = i;
      while (j < paragraph.length && paragraph[j] === ' ') j++;
      tokens.push(paragraph.slice(i, j));
      i = j;
    } else {
      let j = i;
      while (j < paragraph.length) {
        const c = paragraph[j];
        const p = c.codePointAt(0) ?? 0;
        if (c === ' ' || isCjkBreakChar(p)) break;
        j += p > 0xFFFF ? 2 : 1;
      }
      tokens.push(paragraph.slice(i, j));
      i = j;
    }
  }
  let current = '';
  for (const tok of tokens) {
    if (current === '') { current = tok; continue; }
    const candidate = current + tok;
    if (ctx.measureText(candidate).width <= maxWidth) {
      current = candidate;
    } else {
      // Token doesn't fit at the end of the current line — break here.
      // Leading spaces at the start of the next line are dropped (matches
      // Excel: wrapped-continuation lines don't preserve the space that
      // caused the break).
      let nextLead = tok.replace(/^ +/, '');
      if (nextLead === '') nextLead = tok; // all-space token (preserve width on its own line)
      // Apply kinsoku at the boundary: retract trailing code points of the
      // line being closed so it does not end with a 行末禁則 char and the
      // next line does not start with a 行頭禁則 char.
      const lineCps = [...current];
      const retract = kinsokuRetractCount(lineCps, [...nextLead]);
      if (retract > 0) {
        const keep = lineCps.length - retract;
        lines.push(lineCps.slice(0, keep).join(''));
        current = lineCps.slice(keep).join('') + nextLead;
      } else {
        lines.push(current);
        current = nextLead;
      }
    }
  }
  lines.push(current);
  return lines;
}

interface RichSeg {
  text: string;
  font: CellFont;
  width: number; // px
}

interface RichLine {
  segments: RichSeg[];
  maxFontSize: number; // pt (line-height source)
  /** Font family (name) of the run that set `maxFontSize` — the height source
   *  run — so the wrap draw path can floor the single-line height to that
   *  DOCUMENT font's design line box (Meiryo / Sakkal Majalla) via core's
   *  `intendedSingleLinePx`. `null` when the height run named no family. */
  maxFontFamily: string | null;
  /** 0-based index of the LF-delimited paragraph (hard-break region) this line
   *  belongs to. Soft-wrapped continuation lines share their paragraph's index;
   *  it advances only at a hard break. Indexes the per-paragraph bidi base
   *  direction the wrap draw path resolves — UAX#9: a soft wrap does NOT start a
   *  new paragraph, but a hard break does. */
  para: number;
}

/**
 * Layout rich text runs into wrapped lines. Each run is split into words (and
 * CJK characters for granular wrapping). Per-run font is preserved so measurement
 * and drawing use the correct font.
 *
 * Runs are inline and share the cell width (ECMA-376 §18.4.4 r / §18.4.8 si /
 * §18.4.9 sst); wrapText (§18.8.1) breaks at word boundaries (ASCII spaces) and
 * at any CJK code point boundary, and a hard break (LF) starts a new line.
 *
 * An empty value returns `[]` (no fabricated line). This deliberately differs
 * from the plain-text `wrapTextLines`, whose `split('\n')` yields `['']`: an
 * empty cell has no glyphs, so reserving a line for it would only mis-anchor the
 * (non-existent) text.
 */
export function layoutRichTextLines(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  cs: number,
  maxWidth: number,
): RichLine[] {
  const lines: RichLine[] = [];
  let cur: RichSeg[] = [];
  let curW = 0;
  let curMaxSize = 0;
  // Family of the run that currently sets `curMaxSize` on this line — carried so
  // the line-height floor targets the height run's DOCUMENT font.
  let curMaxFamily: string | null = null;
  // Size (pt) of the nearest preceding text run — the height source for a blank
  // line, which has no segment of its own. Mirrors drawShapeText's `lastTextPt`
  // seed (PR #583); falls back to the cell's base font.
  let lastTextPt = baseFont.size;
  // Family paired with `lastTextPt`, so a blank line inherits the nearest
  // preceding text run's family for its own single-line-height floor.
  let lastTextFamily: string | null = baseFont.name;
  // 0-based index of the current LF-delimited paragraph. Advances only at a hard
  // break (not at a soft wrap), so every line records which paragraph it belongs
  // to — the wrap draw path resolves a Context base direction per paragraph.
  let paraIdx = 0;

  // `flush` drops an empty region — used at soft-wrap (kinsoku) breaks, where a
  // line carried wholly to the next line must not leave a blank behind.
  const flush = () => {
    if (cur.length === 0) return;
    lines.push({ segments: cur, maxFontSize: curMaxSize, maxFontFamily: curMaxFamily, para: paraIdx });
    cur = []; curW = 0; curMaxSize = 0; curMaxFamily = null;
  };

  // `flushRegion` emits an empty region as a blank line — used at a hard break
  // (LF) or end-of-value. ECMA-376 §18.8.1 (wrapText): each line of a multi-line
  // cell, including a blank one from consecutive / leading / trailing breaks,
  // reserves one single-line height (the cell analog of PR #583 / docx #582).
  const flushRegion = () => {
    if (cur.length === 0) {
      lines.push({ segments: [], maxFontSize: lastTextPt || DEFAULT_FONT_SIZE, maxFontFamily: lastTextFamily, para: paraIdx });
      return;
    }
    flush();
  };

  const push = (text: string, font: CellFont) => {
    if (!text) return;
    lastTextPt = font.size; // nearest preceding text size, for the next blank line
    lastTextFamily = font.name;
    // Measure at the *draw* font so a super/subscript token reserves its reduced
    // (~65%) glyph width; the segment keeps the run's full size for line height.
    ctx.font = buildFont(vertAlignDrawFont(font), cs);
    // Trailing-space-robust (a wrapped token can end in a space that trimming
    // engines drop; the draw pen advances by the stored width — see the core
    // helper docs).
    const w = measureAdvanceWithTrailingSpace(ctx, text);
    if (cur.length > 0 && curW + w > maxWidth) {
      // Kinsoku at the wrap boundary (ECMA-376 §17.15.1.58–.60): retract
      // trailing code points of the line being closed so it does not end with
      // a 行末禁則 char and the next line (led by `text`) does not start with a
      // 行頭禁則 char. The retracted code points live at the end of the last
      // segment — split that segment (keeping its font), re-measure both parts
      // with the segment's font, and carry the trailing part down to lead the
      // next line.
      const lineCps = cur.flatMap((s) => [...s.text]);
      let retract = kinsokuRetractCount(lineCps, [...text]);
      // UAX#14 LB13: a per-glyph retract would tear a Latin word (e.g. a comma in
      // a separate run overflowing after "system" → "syste" / "m,"). Pull the
      // retract back to the last whitespace so the whole word rides down with the
      // non-starter; CJK boundaries (move char is CJK) are left untouched.
      if (retract > 0) retract = extendLatinWordRetract(lineCps, retract);
      const last = cur[cur.length - 1];
      const lastCps = [...last.text];
      // Only retract within the last segment to preserve each run's font; the
      // single-run CJK case (one segment per line region) is fully covered.
      if (retract > lastCps.length) retract = lastCps.length;
      let carry: RichSeg | null = null;
      if (retract > 0) {
        const keepCps = lastCps.slice(0, lastCps.length - retract);
        const moveCps = lastCps.slice(lastCps.length - retract);
        ctx.font = buildFont(vertAlignDrawFont(last.font), cs);
        if (keepCps.length === 0) {
          // The whole last segment moves down — drop it from the closing line.
          cur.pop();
        } else {
          const keepText = keepCps.join('');
          last.text = keepText;
          // keepText can end at a whitespace (LB13 retract extends to the last
          // space) — measure trailing-space-robustly so it keeps that advance.
          last.width = measureAdvanceWithTrailingSpace(ctx, keepText);
        }
        const moveText = moveCps.join('');
        carry = { text: moveText, font: last.font, width: measureAdvanceWithTrailingSpace(ctx, moveText) };
      }
      flush();
      if (carry) {
        cur.push(carry);
        curW += carry.width;
        if (carry.font.size > curMaxSize) { curMaxSize = carry.font.size; curMaxFamily = carry.font.name; }
      }
      ctx.font = buildFont(vertAlignDrawFont(font), cs); // restore for the incoming token below
    }
    cur.push({ text, font, width: w });
    curW += w;
    if (font.size > curMaxSize) { curMaxSize = font.size; curMaxFamily = font.name; }
  };

  for (const run of runs) {
    const font = applyRunFont(baseFont, run);
    // Tokenize: runs of non-space latin, spaces, or individual CJK chars
    const tokens: string[] = [];
    let i = 0;
    while (i < run.text.length) {
      const ch = run.text[i];
      const cp = ch.codePointAt(0) ?? 0;
      if (cp === 0x000A) {
        // Explicit newline: force break
        tokens.push('\n'); i += 1;
      } else if (isCjkBreakChar(cp)) {
        tokens.push(ch);
        i += cp > 0xFFFF ? 2 : 1;
      } else if (ch === ' ') {
        let j = i;
        while (j < run.text.length && run.text[j] === ' ') j++;
        tokens.push(run.text.slice(i, j));
        i = j;
      } else {
        let j = i;
        while (j < run.text.length) {
          const c = run.text[j];
          const p = c.codePointAt(0) ?? 0;
          if (c === ' ' || c === '\n' || isCjkBreakChar(p)) break;
          j += p > 0xFFFF ? 2 : 1;
        }
        tokens.push(run.text.slice(i, j));
        i = j;
      }
    }
    for (const tok of tokens) {
      // A hard break closes the current paragraph region and opens the next, so
      // the following lines record the new paragraph index. A soft wrap (handled
      // inside `push`) keeps the same index — UAX#9 P1: only a hard break starts
      // a new bidi paragraph.
      if (tok === '\n') { flushRegion(); paraIdx++; }
      else push(tok, font);
    }
  }
  // Trailing region. If the value ended with a break, `cur` is empty but a line
  // was already produced, so a trailing blank line is reserved; a value with no
  // content and no breaks (no segments, no prior line) produces nothing.
  if (cur.length > 0 || lines.length > 0) flushRegion();
  return lines;
}

/** Cell geometry + alignment shared by the rich-text draw helpers (wrap and
 *  non-wrap). `alignH`/`alignV` accept the raw `xf` strings; any value other
 *  than `right`/`center` anchors left, and other than `top`/`center` anchors
 *  bottom (matching the legacy single-line path's handling of `justify` /
 *  `centerContinuous` / etc.). */
export interface RichCellGeom {
  alignH: string;
  alignV: string;
  /** Cell top-left in canvas px (merge span included). */
  cx: number;
  cy: number;
  cellW: number;
  cellH: number;
  /** Left text inset (paddingX + indent) and the symmetric paddings. */
  leftPad: number;
  paddingX: number;
  paddingY: number;
}

/** Underline / strike decoration y for a run on a line, given the line's text
 *  baseline and its `textY`. With a `'top'` baseline `textY` is the line top, so
 *  the underline sits a full text height below it and the strike near mid-height;
 *  `'middle'` / `'bottom'` shift relative to the centre / bottom baseline. The
 *  caller adds any super/subscript `yShift`. One function so the single-line and
 *  multi-line paths place decorations identically. */
function decoYForBaseline(baseline: CanvasTextBaseline, textY: number, rSizePx: number): { underline: number; strike: number } {
  if (baseline === 'middle') return { underline: textY + Math.round(rSizePx * 0.55), strike: textY };
  if (baseline === 'bottom') return { underline: textY + 1, strike: textY - Math.round(rSizePx * 0.35) };
  return { underline: textY + rSizePx + 1, strike: textY + Math.round(rSizePx * 0.5) };
}

/** The draw font for a run/segment: super/subscript renders at ~65% size
 *  (ECMA-376 §18.4.14 vertAlign / ST_VerticalAlignRun §22.9.2.17); everything
 *  else keeps its declared size. The *base* size still governs line height and
 *  the baseline shift — only the glyph (and the width it occupies) shrinks. */
function vertAlignDrawFont(font: CellFont): CellFont {
  return (font.vertAlign === 'superscript' || font.vertAlign === 'subscript')
    ? { ...font, size: font.size * 0.65 } : font;
}

/**
 * Draw a sequence of already-resolved, already-measured rich segments on one
 * line at `textY` under `baseline`, in bidi visual order (UAX#9 rule L2). The
 * caller supplies whether the cell needs the bidi pass and its base direction —
 * a wrapped cell's display lines share one paragraph direction, while hard-break
 * lines are independent paragraphs. Each segment draws with its own font/color,
 * ~65%-size super/subscript baseline shift (§18.4.14 vertAlign / §22.9.2.17), and
 * underline / strike decoration. `startX` is the line's left edge (the caller
 * resolves alignH from the line's measured width). The single shared per-segment
 * drawer for every rich path — non-wrap (via {@link drawRichLine}) and wrap.
 */
export function drawResolvedRichLine(
  ctx: CanvasRenderingContext2D,
  segs: RichSeg[],
  startX: number,
  textY: number,
  baseline: CanvasTextBaseline,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; needBidi?: boolean; baseRtl?: boolean },
): void {
  ctx.textAlign = 'left';
  ctx.textBaseline = baseline;
  const vis = opts.needBidi ? computeLineVisualOrder(segs, opts.baseRtl ?? false) : null;
  const dctx = ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' };
  let x = startX;
  for (let vi = 0; vi < segs.length; vi++) {
    const i = vis ? vis.order[vi] : vi;
    if (vis) { try { dctx.direction = vis.rtl[i] ? 'rtl' : 'ltr'; } catch { /* ignore */ } }
    const seg = segs[i];
    const drawFont = vertAlignDrawFont(seg.font);
    ctx.font = buildFont(drawFont, cs);
    const segColor = opts.fontColor ?? seg.font.color;
    ctx.fillStyle = segColor ? hexToRgba(segColor) : '#000000';
    // Baseline shift for super/subscript, relative to the run's *base* size: up
    // for super, slightly down for sub so each sits at the right vertical band.
    const baseSizePx = vMetricPx(seg.font.size, cs);
    let yShift = 0;
    if (seg.font.vertAlign === 'superscript') yShift = -Math.round(baseSizePx * 0.35);
    else if (seg.font.vertAlign === 'subscript') yShift = Math.round(baseSizePx * 0.10);
    ctx.fillText(seg.text, x, textY + yShift);
    const rSizePx = vMetricPx(drawFont.size, cs);
    if (seg.font.underline || seg.font.strike) {
      const deco = decoYForBaseline(baseline, textY, rSizePx);
      if (seg.font.underline) {
        const stroke = segColor ? hexToRgba(segColor) : '#000000';
        const dbl = seg.font.underlineStyle === 'double' || seg.font.underlineStyle === 'doubleAccounting';
        drawTextDecoLine(ctx, x, x + seg.width, deco.underline + yShift, stroke, dbl, dpr);
      }
      if (seg.font.strike) {
        const syBase = deco.strike + yShift;
        const sy = syBase + crispOffset(syBase, 0.5, dpr);
        ctx.save();
        ctx.strokeStyle = segColor ? hexToRgba(segColor) : '#000000';
        ctx.lineWidth = 0.5;
        ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + seg.width, sy); ctx.stroke();
        ctx.restore();
      }
    }
    x += seg.width;
  }
  // Defensive reset (the per-cell ctx.restore() also restores direction).
  if (vis) { try { dctx.direction = 'ltr'; } catch { /* ignore */ } }
}

/**
 * Draw one already-grouped line of rich-text RUNS at `textY` under `baseline`,
 * positioned horizontally by `geom.alignH` over its measured width. Resolves each
 * run's font (super/subscript at ~65% size, §18.4.14) into a measured segment —
 * the segment keeps the run's *full* size (line height + baseline shift) but its
 * width is measured at the draw font, so a super/subscript run reserves its
 * reduced glyph width — then delegates the painting to {@link drawResolvedRichLine}.
 * Used by the non-wrap rich paths; the wrap path drives `drawResolvedRichLine`
 * directly from its laid-out segments.
 */
function drawRichLine(
  ctx: CanvasRenderingContext2D,
  lineRuns: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number },
  textY: number,
  baseline: CanvasTextBaseline,
): void {
  const { alignH, cx, cellW, leftPad, paddingX } = geom;
  const segs: RichSeg[] = lineRuns.map((r) => {
    const font = applyRunFont(baseFont, r);
    ctx.font = buildFont(vertAlignDrawFont(font), cs);
    // Trailing-space-robust: a rich run can end in a space at an inter-run word
    // boundary, whose advance Firefox/WebKit trim from measureText. The draw pen
    // (drawResolvedRichLine) advances by this stored `width`, so trimming would
    // collide the next run — the core helper restores the space (pass-through on
    // Chrome/skia).
    return { text: r.text, font, width: measureAdvanceWithTrailingSpace(ctx, r.text) };
  });
  const totalWidth = segs.reduce((a, s) => a + s.width, 0);
  let startX: number;
  if (alignH === 'right') startX = cx + cellW - paddingX - totalWidth;
  else if (alignH === 'center') startX = cx + cellW / 2 - totalWidth / 2;
  else startX = cx + leftPad;
  const { needBidi, baseRtl } = resolveCellBidi(opts.readingOrder, segs.map((s) => s.text).join(''));
  drawResolvedRichLine(ctx, segs, startX, textY, baseline, cs, dpr, { fontColor: opts.fontColor, needBidi, baseRtl });
}

/**
 * Draw a single line of rich text (no hard break) with per-run fonts, anchored
 * by the cell's alignV-dependent baseline — `'top'` at the top padding,
 * `'middle'` at the cell centre, `'bottom'` at the bottom padding — exactly how
 * Excel paints a one-line rich cell. Shared by the in-viewport draw path and the
 * off-screen-anchor merge pre-pass so a merged rich cell renders identically
 * whether or not its anchor is scrolled out of view.
 */
function drawSingleLineRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
): void {
  const { alignV, cy, cellH, paddingY } = geom;
  let textY: number;
  let baseline: CanvasTextBaseline;
  if (alignV === 'top') { baseline = 'top'; textY = cy + paddingY; }
  else if (alignV === 'center') { baseline = 'middle'; textY = cy + cellH / 2; }
  else { baseline = 'bottom'; textY = cy + cellH - paddingY; }
  drawRichLine(ctx, runs, baseFont, geom, cs, dpr, opts, textY, baseline);
}

/**
 * Lay out and draw rich text with hard line breaks (LF / Alt+Enter) in a
 * non-wrapped cell. Runs are split at every LF into lines; a blank line from
 * consecutive / leading / trailing breaks reserves one single-line height (the
 * cell analog of PR #585 / docx #582), sized from the nearest preceding text
 * run. Each line is drawn at a `'top'` baseline (matching the wrap rich path)
 * and the whole block is anchored vertically by `alignV` over its summed height.
 */
function drawMultiLineRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
): void {
  const { alignV, cy, cellH, paddingY } = geom;

  // Split runs into lines at LF. A run "A\nB" yields "A" on the current line and
  // "B" on a new one; an empty piece (consecutive / leading / trailing LF) adds
  // no segment but the line still exists, so a blank line is preserved.
  const lineRuns: Run[][] = [[]];
  for (const run of runs) {
    const parts = run.text.split('\n');
    for (let p = 0; p < parts.length; p++) {
      if (p > 0) lineRuns.push([]);
      if (parts[p] !== '') lineRuns[lineRuns.length - 1].push({ ...run, text: parts[p] });
    }
  }

  // Per-line height source (pt) + the family of that height run. A text line uses
  // the max run size on it; a blank line inherits the nearest preceding text
  // run's size AND family — the same seed `layoutRichTextLines` / `drawShapeText`
  // use for blank lines (PR #585). The family drives the single-line-height floor.
  let lastTextPt = baseFont.size;
  let lastTextFamily: string | null = baseFont.name;
  const lineSizes = lineRuns.map((lr) => {
    if (lr.length === 0) return { pt: lastTextPt || DEFAULT_FONT_SIZE, family: lastTextFamily };
    let m = 0;
    let family: string | null = null;
    for (const r of lr) {
      const rf = applyRunFont(baseFont, r);
      if (rf.size > m) { m = rf.size; family = rf.name; }
      lastTextPt = rf.size; // nearest preceding text size, for a following blank line
      lastTextFamily = rf.name;
    }
    return { pt: m, family };
  });
  const lineHeights = lineSizes.map((s) => vMetricPx(s.pt, cs, 1.2, s.family ?? undefined));
  const totalH = lineHeights.reduce((a, b) => a + b, 0);

  let yy: number;
  if (alignV === 'top') yy = cy + paddingY;
  else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
  else yy = cy + cellH - totalH - paddingY;

  for (let li = 0; li < lineRuns.length; li++) {
    const lr = lineRuns[li];
    // A blank line draws nothing but still reserves its height.
    if (lr.length > 0) drawRichLine(ctx, lr, baseFont, geom, cs, dpr, opts, yy, 'top');
    yy += lineHeights[li];
  }
}

/**
 * Draw rich text (mixed-font runs, ECMA-376 §18.4.4 r) in a NON-wrapped cell.
 * A break-free value is one alignV-anchored line ({@link drawSingleLineRichText});
 * a value with a hard break is laid out as multiple lines
 * ({@link drawMultiLineRichText}).
 *
 * §18.8.1 (CT_CellAlignment @wrapText) governs only soft-wrapping; it says
 * nothing about hard breaks. Rendering a literal LF (Alt+Enter, preserved in the
 * run text via §18.4.12 t @xml:space) as a line break even when wrapText is off
 * is undocumented Excel runtime behavior — matched here for parity with the
 * plain-text non-wrap path and the wrap rich-text path (#585).
 */
export function drawNonWrapRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
): void {
  if (runs.some((r) => r.text.includes('\n'))) {
    drawMultiLineRichText(ctx, runs, baseFont, geom, cs, dpr, opts);
  } else {
    drawSingleLineRichText(ctx, runs, baseFont, geom, cs, dpr, opts);
  }
}

/**
 * Lay out and draw WRAP-mode rich text (mixed-font runs, §18.8.1 wrapText on).
 * `layoutRichTextLines` soft-wraps at word / CJK boundaries (and hard breaks),
 * and each resulting line is painted by the shared {@link drawResolvedRichLine}
 * (super/subscript §18.4.14, underline/strike, bidi).
 *
 * Base direction (§18.8.1 readingOrder) is resolved PER LF-delimited paragraph,
 * not once for the whole cell: a soft wrap continues the same bidi paragraph (its
 * display lines share a direction), but a hard break (Alt+Enter LF) starts a new
 * paragraph that, under Context reading order (0/absent), resolves its OWN base
 * direction from its OWN first strong character (UAX#9 P1–P3). Explicit LTR/RTL
 * (1/2) is cell-wide, so only Context varies between paragraphs. Each line carries
 * its paragraph index ({@link RichLine.para}) to pick the matching direction.
 *
 * Drives both the in-viewport draw path and the off-screen-anchor merge pre-pass
 * so a merged wrapped cell renders identically regardless of anchor visibility.
 */
export function drawWrappedRichText(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: CellFont,
  geom: RichCellGeom,
  cs: number,
  dpr: number,
  opts: { fontColor?: string | null; readingOrder?: number } = {},
): void {
  const { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY } = geom;
  const rLines = layoutRichTextLines(ctx, runs, baseFont, cs, cellW - leftPad - paddingX);
  const totalH = rLines.reduce((s, l) => s + vMetricPx(l.maxFontSize, cs, 1.2, l.maxFontFamily ?? undefined), 0);
  let yy: number;
  if (alignV === 'top') yy = cy + paddingY;
  else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
  else yy = cy + cellH - totalH - paddingY;
  // Resolve the bidi base direction for each LF-delimited paragraph from its own
  // text (Context = first-strong; explicit 1/2 = cell-wide). `line.para` indexes
  // this, so soft-wrapped lines share their paragraph's direction while a hard
  // break gets its own. `resolveCellBidi` gates each so pure-LTR paragraphs keep
  // the exact pre-bidi path — the same gate the non-wrap path uses per line.
  const paraTexts = runs.map((r) => r.text).join('').split('\n');
  const paraBidi = paraTexts.map((t) => resolveCellBidi(opts.readingOrder, t));
  for (const line of rLines) {
    const totalW = line.segments.reduce((s, seg) => s + seg.width, 0);
    let xx: number;
    if (alignH === 'right') xx = cx + cellW - paddingX - totalW;
    else if (alignH === 'center') xx = cx + cellW / 2 - totalW / 2;
    else xx = cx + leftPad;
    const { needBidi, baseRtl } = paraBidi[line.para];
    drawResolvedRichLine(ctx, line.segments, xx, yy, 'top', cs, dpr, { fontColor: opts.fontColor, needBidi, baseRtl });
    yy += vMetricPx(line.maxFontSize, cs, 1.2, line.maxFontFamily ?? undefined);
  }
}

function colToLetter(col: number): string {
  let result = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

interface RenderContext {
  worksheet: Worksheet;
  styles: Styles;
  cellMap: Map<string, Cell>;
  mergeAnchorMap: Map<string, { totalW: number; totalH: number; right: number; bottom: number }>;
  mergeSkipSet: Set<string>;
  cfContext: CfContext;
  colWidths: number[];
  rowHeights: number[];
  frozenColWidths: number[];
  frozenRowHeights: number[];
  frozenW: number;
  frozenH: number;
  startRow: number;
  startCol: number;
  cs: number;
  dpr: number;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
  /** row:col keys for cells that carry a comment; renderer draws a small
   *  red triangle in the top-right corner (ECMA-376 §18.7.3 commentList). */
  commentCells: Set<string>;
  /** row:col → table-style overlay (bold header, banded rows, borders). */
  tableStyleMap: Map<string, TableCellStyle>;
  /** row:col → render-ready SparklineModel for cells that host an
   *  `x14:sparkline`. Built once at viewport start by flattening the
   *  parser's SparklineGroup + per-cell Sparkline pair. */
  sparklineMap: Map<string, SparklineModel>;
  /** Max Digit Width resolved for the worksheet's Normal-style font
   *  (ECMA-376 §18.3.1.13). Used by `colWidthToPx` to convert character-
   *  unit column widths into pixels. */
  mdw: number;
  onTextRun?: (info: XlsxTextRunInfo) => void;
  /** Sheet-level right-to-left grid mirror (ECMA-376 §18.3.1.87
   *  `<sheetView rightToLeft>`). When true the whole grid is mirrored:
   *  column A sits at the right edge, columns flow right-to-left, and the
   *  row-number header strip sits on the right of the canvas. Implemented
   *  by remapping each cell's canvas x to `canvasW - cx - cellW` (mirror
   *  about the cell-area band) — glyphs themselves are NOT flipped, and
   *  cell-level left/right alignment stays physical (Excel behavior). */
  rtl: boolean;
  /** Logical canvas width (device-independent px). Needed by the RTL
   *  mirror; identical to the value the top-level render fn computes. */
  canvasW: number;
}

// ────────────────────────────────────────────────────────────────
// Icon Set drawing
// ────────────────────────────────────────────────────────────────
const ICON_COLORS_3 = ['#FF0000', '#FFFF00', '#00B050'];
const ICON_COLORS_4 = ['#FF0000', '#FF6600', '#FFFF00', '#00B050'];
const ICON_COLORS_5 = ['#FF0000', '#FF6600', '#FFFF00', '#92D050', '#00B050'];

function drawCfIcon(ctx: CanvasRenderingContext2D, name: string, index: number, x: number, y: number, sz: number): void {
  if (name === 'NoIcons') return;
  const safeName = name || '3TrafficLights1';
  const nIcons = parseInt(safeName[0]) || 3;
  const palette = nIcons === 5 ? ICON_COLORS_5 : nIcons === 4 ? ICON_COLORS_4 : ICON_COLORS_3;
  const color = palette[Math.max(0, Math.min(index, palette.length - 1))];
  ctx.save();
  ctx.fillStyle = color;
  if (safeName.includes('Arrow')) {
    const half = sz / 2;
    ctx.beginPath();
    if (index === nIcons - 1) {
      ctx.moveTo(x + half, y); ctx.lineTo(x + sz, y + sz); ctx.lineTo(x, y + sz);
    } else if (index === 0) {
      ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x + half, y + sz);
    } else {
      ctx.moveTo(x, y + sz * 0.3); ctx.lineTo(x + sz, y + half); ctx.lineTo(x, y + sz * 0.7);
    }
    ctx.closePath();
    ctx.fill();
  } else if (safeName.includes('Flag')) {
    ctx.beginPath();
    ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x, y + sz);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.arc(x + sz / 2, y + sz / 2, sz / 2, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}

function drawAutoFilterArrow(ctx: CanvasRenderingContext2D, cx: number, cy: number, cw: number, ch: number): void {
  const sz = Math.max(6, Math.round(Math.min(cw, ch) * 0.45));
  const x = cx + cw - sz - 1;
  const y = cy + ch - sz - 1;
  ctx.save();
  ctx.fillStyle = '#D0D0D0';
  ctx.fillRect(x, y, sz, sz);
  ctx.fillStyle = '#444444';
  const tri = sz * 0.55;
  const tx = x + (sz - tri) / 2;
  const ty = y + (sz - tri * 0.5) / 2;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(tx + tri, ty);
  ctx.lineTo(tx + tri / 2, ty + tri * 0.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Excel Table style overlays (ECMA-376 §18.5 / §18.8.83)
// ────────────────────────────────────────────────────────────────
// Two distinct paths:
//
// 1. CUSTOM styles (defined in the file's `<tableStyles>` block, §18.5.1.2):
//    rendered strictly from the dxfs of their declared `<tableStyleElement>`s.
//    A custom style contributes ONLY what its elements define — if none carry
//    a border, Excel draws no table-level border (the visible structure comes
//    from theme borders baked into each cell `xf`, which we render separately).
//    `isCustom` cells therefore never get accent synthesis.
//
// 2. BUILT-IN style names (`TableStyleLight18`, …) whose definitions are NOT
//    in the file. We don't yet ship the preset catalog, so we approximate them
//    from a single "accent" color: bold header + banded fills + horizontal
//    rules, so those files render with visible structure rather than blank
//    ranges. This is an approximation, kept until a real catalog ships
//    (post-1.0).
export interface TableCellStyle {
  accent: string;
  /** `true` for a custom `<tableStyle>` — disables all accent approximation. */
  isCustom: boolean;
  isHeader: boolean;
  isTotals: boolean;
  /** `true` when this is a banded data row that should get the stripe fill. */
  isBanded: boolean;
  isFirstCol: boolean;
  isLastCol: boolean;
  isTopEdge: boolean;
  isBottomEdge: boolean;
  /** Dxf for the whole-table element of a custom `<tableStyle>`
   *  (ECMA-376 §18.8.83). Border/fill apply to every cell as a base layer. */
  wholeTableDxf?: number;
  /** Dxf for the header-row element of a custom `<tableStyle>`. Provides
   *  header fill, font color/weight, and vertical separators. */
  headerRowDxf?: number;
  /** Dxf for the total-row element. */
  totalRowDxf?: number;
  /** Dxf for the first-column element (when `showFirstColumn`). */
  firstColumnDxf?: number;
  /** Dxf for the last-column element (when `showLastColumn`). */
  lastColumnDxf?: number;
  /** Dxf for the stripe (band1=odd, band2=even) that applies to this row when
   *  `showRowStripes` is set; undefined when this row is not a stripe. */
  stripeDxf?: number;
}

function buildTableStyleMap(worksheet: Worksheet): Map<string, TableCellStyle> {
  const map = new Map<string, TableCellStyle>();
  for (const t of worksheet.tables ?? []) {
    // ECMA-376 §18.5.1.4: empty styleName means the table has "None" style —
    // no visual table formatting overlay should be applied. Cell xf borders
    // and fills are still rendered as defined (per §18.8.45); only the
    // table-style overlay (banded fills, separator rules, etc.) is skipped.
    if (!t.styleName) continue;
    const { top, bottom, left, right } = t.range;
    const accent = t.accentColor || '#808080';
    const isCustom = !!t.isCustom;
    const hdr = Math.max(0, t.headerRowCount ?? 1);
    const tot = Math.max(0, t.totalsRowCount ?? 0);
    const headerEnd = top + hdr - 1;
    const totalsStart = bottom - tot + 1;
    for (let r = top; r <= bottom; r++) {
      const isHeader = hdr > 0 && r <= headerEnd;
      const isTotals = tot > 0 && r >= totalsStart;
      const dataIdx = (!isHeader && !isTotals) ? (r - headerEnd - 1) : -1;
      // Row banding stripes alternate band1 (odd) / band2 (even) over the data
      // region (§18.18.93). For built-in approximation we only paint the odd
      // stripe via `isBanded`; for custom styles we pick the matching dxf.
      const isStripeRow = t.showRowStripes && dataIdx >= 0;
      const stripeDxf = isStripeRow
        ? (dataIdx % 2 === 1 ? t.band1HorizontalDxf : t.band2HorizontalDxf)
        : undefined;
      for (let c = left; c <= right; c++) {
        map.set(`${r}:${c}`, {
          accent,
          isCustom,
          isHeader,
          isTotals,
          isBanded: t.showRowStripes && dataIdx >= 0 && dataIdx % 2 === 1,
          isFirstCol: t.showFirstColumn && c === left,
          isLastCol: t.showLastColumn && c === right,
          isTopEdge: r === top,
          isBottomEdge: r === bottom,
          wholeTableDxf: t.wholeTableDxf,
          headerRowDxf: t.headerRowDxf,
          totalRowDxf: t.totalRowDxf,
          firstColumnDxf: t.firstColumnDxf,
          lastColumnDxf: t.lastColumnDxf,
          stripeDxf,
        });
      }
    }
  }
  return map;
}

/**
 * Result of resolving a table cell's overlay *border* (ECMA-376 §18.8.83).
 *  - `none`   — draw nothing extra (custom style with no border dxf, or a
 *               built-in cell that has no rule on this edge).
 *  - `dxf`    — draw `border` exactly as the resolved dxf defines it.
 *  - `accent` — built-in approximation only: a synthesized accent-colored rule
 *               under the cell (plus the table top edge when `topEdge`).
 */
export type TableOverlayBorder =
  | { kind: 'none' }
  | { kind: 'dxf'; border: Border }
  | { kind: 'accent'; color: string; lineWidth: number; topEdge: boolean };

/**
 * Decide the table-style overlay border for one cell.
 *
 * `dxfWhole` / `dxfHeader` are the already-resolved dxfs for the cell's
 * wholeTable / headerRow elements (undefined when the element is absent).
 * `colIndex` is the cell's column index *within the table* (0 = first column),
 * used to draw the table's outer-left edge only on the leftmost column.
 *
 * Custom styles (`ts.isCustom`) draw only from these dxfs and never synthesize
 * accent rules: when no border dxf is present they contribute nothing. Built-in
 * style names keep the accent approximation.
 */
export function tableOverlayBorder(
  ts: TableCellStyle,
  dxfWhole: Dxf | undefined,
  dxfHeader: Dxf | undefined,
  colIndex: number,
): TableOverlayBorder {
  const horiz = dxfWhole?.border?.horizontal;
  const wtTop = dxfWhole?.border?.top;
  const wtBot = dxfWhole?.border?.bottom;
  const wtLeft = dxfWhole?.border?.left;
  const wtRight = dxfWhole?.border?.right;
  const hdrBot = dxfHeader?.border?.bottom;
  const hdrTop = dxfHeader?.border?.top;
  const hasDxfBorder = !!(horiz || wtTop || wtBot || wtLeft || wtRight || hdrBot || hdrTop);

  if (hasDxfBorder) {
    // Compose per-edge from the table-style hierarchy
    // (wholeTable < headerRow), §18.8.83. The inner `horizontal` rule fills
    // top/bottom for interior rows; outer edges come from the table extents.
    const overlay: Border = { left: null, right: null, top: null, bottom: null };
    if (ts.isTopEdge) overlay.top = wtTop ?? null;
    else if (horiz) overlay.top = horiz;
    if (ts.isHeader && hdrBot) overlay.bottom = hdrBot;
    else if (ts.isBottomEdge) overlay.bottom = wtBot ?? null;
    else if (horiz) overlay.bottom = horiz;
    if (ts.isFirstCol || colIndex === 0) overlay.left = wtLeft ?? null;
    if (ts.isLastCol) overlay.right = wtRight ?? null;
    return { kind: 'dxf', border: overlay };
  }

  // No border dxf. A custom style contributes nothing here (§18.5.1.2):
  // Excel draws no table-level border, only the theme borders baked into the
  // cell xf. Built-in names fall through to the accent approximation.
  if (ts.isCustom) return { kind: 'none' };
  return { kind: 'accent', color: ts.accent, lineWidth: ts.isHeader ? 1.5 : 1, topEdge: ts.isTopEdge };
}

/** Flatten the worksheet's parsed `sparklineGroups` into a per-cell render
 *  model. Each Sparkline inherits its group's formatting; min/max are
 *  computed from the values when the group's `*AxisType` is `individual`,
 *  shared across the group when `group`, or taken from `manualMin/Max` when
 *  `custom`. The renderer can then look up `row:col` and call core's
 *  `renderSparkline` without further work. */
function buildSparklineMap(worksheet: Worksheet): Map<string, SparklineModel> {
  const map = new Map<string, SparklineModel>();
  for (const g of worksheet.sparklineGroups ?? []) {
    // Group-wide min/max if needed.
    let groupMin = Infinity, groupMax = -Infinity;
    if (g.minAxisType === 'group' || g.maxAxisType === 'group') {
      for (const sl of g.sparklines) {
        for (const v of sl.values) {
          if (typeof v === 'number') {
            if (v < groupMin) groupMin = v;
            if (v > groupMax) groupMax = v;
          }
        }
      }
      if (!isFinite(groupMin) || !isFinite(groupMax)) {
        groupMin = 0; groupMax = 1;
      }
    }
    for (const sl of g.sparklines) {
      const numeric = sl.values.filter((v): v is number => typeof v === 'number');
      const indMin = numeric.length ? Math.min(...numeric) : 0;
      const indMax = numeric.length ? Math.max(...numeric) : 1;
      const min = g.minAxisType === 'custom' && typeof g.manualMin === 'number'
        ? g.manualMin
        : g.minAxisType === 'group' ? groupMin : indMin;
      const max = g.maxAxisType === 'custom' && typeof g.manualMax === 'number'
        ? g.manualMax
        : g.maxAxisType === 'group' ? groupMax : indMax;
      map.set(`${sl.row}:${sl.col}`, {
        kind: g.kind,
        values: sl.values,
        min,
        max,
        displayEmptyCellsAs: (g.displayEmptyCellsAs === 'zero' || g.displayEmptyCellsAs === 'span')
          ? g.displayEmptyCellsAs
          : 'gap',
        displayXAxis: g.displayXAxis,
        lineWeight: g.lineWeight,
        markers: g.markers,
        high: g.high,
        low: g.low,
        first: g.first,
        last: g.last,
        negative: g.negative,
        colorSeries: g.colorSeries,
        colorNegative: g.colorNegative,
        colorAxis: g.colorAxis,
        colorMarkers: g.colorMarkers,
        colorFirst: g.colorFirst,
        colorLast: g.colorLast,
        colorHigh: g.colorHigh,
        colorLow: g.colorLow,
      });
    }
  }
  return map;
}

function stripeColorFor(accent: string): string {
  // Light tint of the accent — mimics TableStyleLight* banded rows.
  const hex = accent.replace('#', '');
  if (hex.length < 6) return '#F2F2F2';
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const mix = (ch: number) => Math.round(ch * 0.2 + 255 * 0.8);
  const toHex = (v: number) => v.toString(16).padStart(2, '0').toUpperCase();
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

// ────────────────────────────────────────────────────────────────
// Render one rectangular region of cells
// ────────────────────────────────────────────────────────────────
function renderQuadrant(
  ctx: CanvasRenderingContext2D,
  rc: RenderContext,
  startRow: number, startCol: number,
  colWidths: number[], rowHeights: number[],
  pixOffsetX: number, pixOffsetY: number,
  originX: number, originY: number,
  clipX: number, clipY: number, clipW: number, clipH: number,
): void {
  if (clipW <= 0 || clipH <= 0) return;

  const { styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext, cs, dpr } = rc;
  const numCols = colWidths.length;
  const numRows = rowHeights.length;

  // RTL grid mirror (ECMA-376 §18.3.1.87). Maps a left-anchored cell rect
  // [x, x+w] to its mirror [canvasW - x - w, canvasW - x] within the
  // cell-area band. Because the LTR cell area is [hw, canvasW] and the RTL
  // cell area is [0, canvasW - hw] (header strip moves to the right), the
  // transform for any cell/quadrant collapses to `canvasW - x - w`. Applied
  // to every column x (so fills, borders, gridlines, text positions, merge
  // spans and the clip rect all mirror) while leaving glyphs un-flipped and
  // cell-level left/right alignment physical.
  const mirrorX = (x: number, w: number): number => rc.rtl ? rtlMirrorX(x, w, rc.canvasW) : x;

  // Canvas x for each column
  const colXs: number[] = [];
  let x = -pixOffsetX;
  for (let ci = 0; ci < numCols; ci++) {
    colXs.push(x);
    x += colWidths[ci];
  }

  // Canvas y for each row
  const rowYs: number[] = [];
  let y = -pixOffsetY;
  for (let ri = 0; ri < numRows; ri++) {
    rowYs.push(y);
    y += rowHeights[ri];
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(mirrorX(clipX, clipW), clipY, clipW, clipH);
  ctx.clip();

  // Deferred text drawing. Excel renders cell text on top of *all* cell
  // backgrounds, so a left-aligned overflow into an adjacent empty cell with
  // a white fill stays visible. If we drew fill+text per cell in one pass,
  // the next cell's fill would overpaint the previous cell's overflow.
  const textTasks: Array<() => void> = [];

  // Deferred merged-cell border drawing. Cell gridlines (#d0d0d0) are stroked
  // interleaved with cell borders inside the grid loop, at each cell's right +
  // bottom edge. For a 1×1 cell that ordering is harmless (within a row the
  // left neighbour is visited — and draws its right gridline — before the cell
  // strokes its own left border over it). A merged anchor, however, draws its
  // border ONCE at full span height/width during the anchor's row, while the
  // cells in the column to the left of the covered rows are visited in LATER
  // rows and stroke their right gridline over the merge's already-drawn left
  // border — shaving a device column off it (sample-30 B14:B23 left edge read
  // thinner/greyer below the anchor). Excel treats an explicit cell border as
  // replacing the gridline on that edge, so the merge perimeter must sit above
  // every gridline. Defer it to a pass flushed after the whole grid loop (but
  // before text, which stays on top — same z-order as the per-cell path).
  const mergeBorderTasks: Array<() => void> = [];

  // Deferred NON-merged cell borders + table-style overlays. Two-pass painting:
  // every cell fill (and the gray sheet gridlines) is laid down in the main
  // loop (pass 1); all styled cell borders run here (pass 2), AFTER every fill.
  // This removes the latent "a neighbour's fill over-paints a shared border"
  // bug structurally — once borders are deferred past all fills, no fill can
  // erase them — so the old `invertedTop`/`invertedLeft` inherit-and-repair
  // mechanism is gone (only genuine conflict resolution via `pickStrongerEdge`
  // remains). Mirrors docx PR #545 and the pptx renderer's two-pass order.
  const borderTasks: Array<() => void> = [];

  // Pre-pass: merge cells whose anchor lies outside this viewport quadrant but whose
  // span overlaps it (e.g. scrolled past the anchor row/col, or the anchor is in a
  // frozen quadrant while we are rendering the scrollable quadrant).
  for (const mc of rc.worksheet.mergeCells ?? []) {
    const aRow = mc.top, aCol = mc.left;
    // If anchor is within the main loop range, skip — handled normally below.
    if (aRow >= startRow && aRow < startRow + numRows &&
        aCol >= startCol && aCol < startCol + numCols) continue;
    // Skip if merge span has no overlap with this viewport.
    if (mc.bottom < startRow || mc.top >= startRow + numRows) continue;
    if (mc.right  < startCol || mc.left >= startCol + numCols) continue;

    const info = rc.mergeAnchorMap.get(`${aRow}:${aCol}`);
    if (!info) continue;

    // Canvas X of anchor col (may be negative = off-screen to the left).
    let aCx: number;
    if (aCol >= startCol) {
      aCx = originX + colXs[aCol - startCol];
    } else {
      let dx = 0;
      for (let c = aCol; c < startCol; c++) {
        dx += Math.round(colWidthToPx(rc.worksheet.colWidths[c] ?? rc.worksheet.defaultColWidth, rc.mdw) * cs);
      }
      aCx = originX - pixOffsetX - dx;
    }
    // Canvas Y of anchor row (may be negative = off-screen above).
    let aCy: number;
    if (aRow >= startRow) {
      aCy = originY + rowYs[aRow - startRow];
    } else {
      let dy = 0;
      for (let r = aRow; r < startRow; r++) {
        dy += Math.round(rowHeightToPx(rc.worksheet.rowHeights[r] ?? rc.worksheet.defaultRowHeight) * cs);
      }
      aCy = originY - pixOffsetY - dy;
    }

    const cW = info.totalW, cH = info.totalH;
    aCx = mirrorX(aCx, cW);
    const key = `${aRow}:${aCol}`;
    const cell = rc.cellMap.get(key);
    const { font, fill, border, xf } = resolveXf(styles, cell?.styleIndex ?? 0);
    const cf = evaluateCf(cell, aRow, aCol, cfContext, styles.dxfs ?? []);
    const effectiveFill = cf.fill ?? fill;

    paintCellPatternFill(ctx, effectiveFill, aCx, aCy, cW, cH);
    if (cf.dataBar && cf.dataBar.ratio > 0) {
      const bInset = 2;
      const bW = Math.max(0, (cW - bInset * 2) * cf.dataBar.ratio);
      fillDataBar(ctx, cf.dataBar.color, aCx + bInset, aCy + bInset, bW, cH - bInset * 2, cf.dataBar.gradient);
    }
    // Defer this off-screen-anchored merge's border too: its in-viewport span
    // is exposed to the same gridline overpaint (left-column cells in the
    // covered rows stroke their right gridline over it in the main loop). Note
    // this pre-pass draws its own text *inline* below (the main loop defers
    // text via `textTasks`), so the deferred border ends up above this cell's
    // text rather than below it — visible only if a *thick* border grazed the
    // padding-inset, clipped text of an off-screen-anchored merge. Routing this
    // pre-pass text through `textTasks` for full gridline<border<text parity is
    // a follow-up; the current order is visually inert.
    const mergedBorder = resolveMergeBorder(border, aRow, aCol, info.right, info.bottom, rc.cellMap, styles);
    const preBorder = mergeBorders(mergedBorder, cf.border);
    mergeBorderTasks.push(() => renderBorder(ctx, preBorder, aCx, aCy, cW, cH, dpr));

    if (!cell) continue;
    const text = formatCellValue(cell, styles, cf.numFmt, rc.worksheet.date1904);
    if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

    const effectiveBold = font.bold || !!cf.fontBold;
    const effectiveItalic = font.italic || !!cf.fontItalic;
    const effectiveUnderline = font.underline || !!cf.fontUnderline;
    const effectiveStrike = font.strike || !!cf.fontStrike;
    const fontForDraw: CellFont = (
      effectiveBold !== font.bold || effectiveItalic !== font.italic ||
      effectiveUnderline !== font.underline || effectiveStrike !== font.strike
    ) ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
      : font;
    ctx.font = buildFont(fontForDraw, cs);
    const hyperlinkUrl = rc.hyperlinkMap.get(key);
    const textColor = hyperlinkUrl ? '#0563C1' : (cf.fontColor ?? font.color);
    ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

    const paddingX = 3, paddingY = 2;
    const isNumeric = cell.value.type === 'number';
    const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
    const alignV = xf.alignV ?? 'bottom';
    // Indent: ECMA-376 §18.8.1 alignment@indent — one level indents by 3
    // character widths (MDW) of the workbook's normal-style font.
    const indentPx = xf.indent ? Math.round(xf.indent * 3 * rc.mdw) : 0;
    const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0);

    ctx.save();
    ctx.beginPath();
    ctx.rect(aCx, aCy, cW, cH);
    ctx.clip();

    let textX: number;
    if (alignH === 'right') { textX = aCx + cW - paddingX; ctx.textAlign = 'right'; }
    else if (alignH === 'center') { textX = aCx + cW / 2; ctx.textAlign = 'center'; }
    else { textX = aCx + leftPad; ctx.textAlign = 'left'; }

    // Text layout mirrors the main loop's merged-anchor path (see the
    // `xf.wrapText` branches below): wrapped cells split into lines via
    // wrapTextLines / layoutRichTextLines and flow vertically per alignV.
    // Using the same code keeps the off-screen-anchor pre-pass and the
    // in-viewport-anchor path identical, so a merged cell renders the same
    // text whether or not its top-left anchor cell is scrolled out of view.
    const runs = cell.value.type === 'text' ? cell.value.runs : undefined;
    const hasRichText = runs && runs.length > 0;

    if (xf.wrapText && hasRichText) {
      // Same helper as the in-viewport wrap path so an off-screen-anchored merge
      // renders identical wrapped rich text — per-run fonts, super/subscript,
      // underline/strike, and bidi (previously this pre-pass drew only plain
      // per-segment fonts).
      drawWrappedRichText(
        ctx, runs, fontForDraw,
        { alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY },
        cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder },
      );
    } else if (xf.wrapText) {
      const lines = wrapTextLines(ctx, text, cW - leftPad - paddingX);
      const lineH = vMetricPx(font.size, cs, 1.2, font.name ?? undefined);
      const totalTextH = lines.length * lineH;
      let startY: number;
      if (alignV === 'top') startY = aCy + paddingY;
      else if (alignV === 'center') startY = aCy + (cH - totalTextH) / 2;
      else startY = aCy + cH - totalTextH - paddingY;
      ctx.textBaseline = 'top';
      for (let li = 0; li < lines.length; li++) {
        ctx.fillText(lines[li], textX, startY + li * lineH);
      }
    } else if (hasRichText) {
      // Non-wrap rich text — same helper as the in-viewport path so an
      // off-screen-anchored merge renders identical per-run text (single line, or
      // multiple lines on a hard break) instead of joined base-font text.
      drawNonWrapRichText(
        ctx, runs, fontForDraw,
        { alignH, alignV, cx: aCx, cy: aCy, cellW: cW, cellH: cH, leftPad, paddingX, paddingY },
        cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder },
      );
    } else {
      let textY: number;
      if (alignV === 'top') { ctx.textBaseline = 'top'; textY = aCy + paddingY; }
      else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = aCy + cH / 2; }
      else { ctx.textBaseline = 'bottom'; textY = aCy + cH - paddingY; }
      ctx.fillText(text, textX, textY);
    }
    ctx.restore();
  }

  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = startRow + ri;
    const cy = originY + rowYs[ri];
    const ch = rowHeights[ri];
    if (cy + ch <= clipY || cy >= clipY + clipH) continue;

    // Pre-compute centerContinuous ranges in this row. ECMA-376 §18.18.40:
    // a contiguous run of cells whose alignment is `centerContinuous` is
    // treated as a single visual span — Excel hides the default gridlines
    // *and* the explicit cell borders inside the run, so the run reads as
    // one merged-looking span with only the outer perimeter visible.
    //
    // A run is bounded by either (a) a non-centerContinuous cell, or
    // (b) another centerContinuous cell that itself carries a value —
    // the spec says spanned cells reference the same style id while only
    // the anchor holds the displayed value. Two anchors in a row therefore
    // mean two adjacent runs, with the border between them visible.
    const suppressRightGridCol = new Set<number>();
    const suppressLeftGridCol = new Set<number>();
    let runStart = -1;
    const closeRun = (endExclusive: number) => {
      if (runStart >= 0 && endExclusive - runStart >= 2) {
        for (let k = runStart; k < endExclusive - 1; k++) suppressRightGridCol.add(k);
        for (let k = runStart + 1; k < endExclusive; k++) suppressLeftGridCol.add(k);
      }
      runStart = -1;
    };
    for (let ci = 0; ci <= numCols; ci++) {
      let isCC = false;
      let hasValue = false;
      if (ci < numCols) {
        const ckey = `${rowIndex}:${startCol + ci}`;
        if (!mergeSkipSet.has(ckey) && !mergeAnchorMap.has(ckey)) {
          const c = cellMap.get(ckey);
          const cXf = resolveXf(styles, c?.styleIndex ?? 0).xf;
          isCC = cXf.alignH === 'centerContinuous';
          hasValue = !!(c && c.value && c.value.type !== 'empty');
        }
      }
      if (!isCC) {
        closeRun(ci);
      } else if (hasValue && runStart >= 0 && ci > runStart) {
        // New anchor: close previous run [runStart, ci) and start a fresh one.
        closeRun(ci);
        runStart = ci;
      } else if (runStart < 0) {
        runStart = ci;
      }
    }

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = startCol + ci;
      const ltrCx = originX + colXs[ci];
      const cw = colWidths[ci];
      // Cull against the (un-mirrored) clip band — visibility is preserved
      // by the mirror, so the LTR test is correct for both directions.
      if (ltrCx + cw <= clipX || ltrCx >= clipX + clipW) continue;

      const key = `${rowIndex}:${colIndex}`;
      if (mergeSkipSet.has(key)) continue;

      const mergeInfo = mergeAnchorMap.get(key);
      const cellW = mergeInfo ? mergeInfo.totalW : cw;
      const cellH = mergeInfo ? mergeInfo.totalH : ch;
      // Mirror the cell's canvas x for RTL. Uses the *full* drawn width
      // (merge span included) so the rect lands at the correct mirror band.
      const cx = mirrorX(ltrCx, cellW);

      const cell = cellMap.get(key);
      const styleIndex = cell?.styleIndex ?? 0;
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);
      const cf = evaluateCf(cell, rowIndex, colIndex, cfContext, styles.dxfs ?? []);
      const effectiveFill = cf.fill ?? fill;
      const tableStyle = rc.tableStyleMap.get(key);
      // Custom `<tableStyle>` dxfs (ECMA-376 §18.8.83). When present, they
      // drive fills / font color and borders strictly from the declared
      // elements — no accent fallback for custom styles.
      const dxfList = styles.dxfs ?? [];
      const dxfAt = (i?: number) => (i != null ? dxfList[i] : undefined);
      const tsDxfWhole = dxfAt(tableStyle?.wholeTableDxf);
      const tsDxfHeader = dxfAt(tableStyle?.headerRowDxf);
      const tsDxfTotal = dxfAt(tableStyle?.totalRowDxf);
      const tsDxfFirstCol = dxfAt(tableStyle?.firstColumnDxf);
      const tsDxfLastCol = dxfAt(tableStyle?.lastColumnDxf);
      const tsDxfStripe = dxfAt(tableStyle?.stripeDxf);
      // Per-cell table fill resolved from the element hierarchy
      // (§18.8.83: wholeTable < band < column < header/total). The later a
      // layer appears the higher its precedence; we pick the most specific
      // fill that defines a fgColor.
      const tableFillDxf =
        (tableStyle?.isHeader && tsDxfHeader?.fill?.fgColor) ? tsDxfHeader :
        (tableStyle?.isTotals && tsDxfTotal?.fill?.fgColor) ? tsDxfTotal :
        (tableStyle?.isLastCol && tsDxfLastCol?.fill?.fgColor) ? tsDxfLastCol :
        (tableStyle?.isFirstCol && tsDxfFirstCol?.fill?.fgColor) ? tsDxfFirstCol :
        (tsDxfStripe?.fill?.fgColor) ? tsDxfStripe :
        (!tableStyle?.isHeader && !tableStyle?.isTotals && tsDxfWhole?.fill?.fgColor) ? tsDxfWhole :
        undefined;

      // Background fill (base or CF override). ECMA-376 §18.8.22 ST_PatternType.
      // - solid/gray*: blend fgColor with bgColor at the pattern's fg coverage.
      // - directional hatches (dark/light Horizontal/Vertical/Down/Up/Grid/
      //   Trellis): render via a small repeating tile using createPattern so
      //   the hatch actually shows, rather than approximating as a blend.
      if (paintCellPatternFill(ctx, effectiveFill, cx, cy, cellW, cellH)) {
        // own fill painted; tableStyle fallbacks intentionally skipped
      } else if (tableStyle && tableFillDxf?.fill?.fgColor) {
        // Custom or built-in: a resolved table-element dxf fill wins.
        ctx.fillStyle = hexToRgba(tableFillDxf.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && !tableStyle.isCustom && tableStyle.isBanded) {
        // Accent-tint banding is an approximation for built-in style names
        // only. Custom styles with no stripe dxf get no banding fill (Excel
        // draws none — §18.5.1.2).
        ctx.fillStyle = stripeColorFor(tableStyle.accent);
        ctx.fillRect(cx, cy, cellW, cellH);
      }

      // Comment indicator triangle — drawn above fill but below borders so
      // borders still read cleanly around the cell edge.
      if (rc.commentCells.has(key)) {
        drawCommentMarker(ctx, cx, cy, cellW, cellH);
      }

      // DataBar (drawn inside the cell, left-anchored). Excel 2010+ renders
      // these with a horizontal gradient by default.
      if (cf.dataBar && cf.dataBar.ratio > 0) {
        const barInset = 2;
        const barW = Math.max(0, (cellW - barInset * 2) * cf.dataBar.ratio);
        fillDataBar(ctx, cf.dataBar.color, cx + barInset, cy + barInset, barW, cellH - barInset * 2, cf.dataBar.gradient);
      }

      // Sparkline (Office 2010 x14:sparklineGroup). Drawn after the cell
      // background but before borders / text so borders frame the sparkline
      // and any cell text overlays it (matches Excel's z-order, and lets
      // a label like "Trend" share the same cell as the sparkline).
      const sparkModel = rc.sparklineMap.get(key);
      if (sparkModel) {
        renderSparkline(ctx, { x: cx, y: cy, w: cellW, h: cellH }, sparkModel);
      }

      // Grid lines – draw only right + bottom edges once per cell (avoids double-drawing at
      // shared cell boundaries). crispOffset snaps each line onto the device
      // pixel grid so we get a crisp 1-device-pixel result at any dpr (a fixed
      // 0.5/dpr offset blurs at even device widths, e.g. dpr=3).
      // Skipped when the sheet has `<sheetView showGridLines="0">` (View →
      // Gridlines unchecked; ECMA-376 §18.3.1.83).
      if (rc.worksheet.showGridlines !== false) {
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.beginPath();
        if (!suppressRightGridCol.has(ci)) {
          const rx = cx + cellW + crispOffset(cx + cellW, 0.5, dpr);  // right edge
          ctx.moveTo(rx, cy);
          ctx.lineTo(rx, cy + cellH);
        }
        const by = cy + cellH + crispOffset(cy + cellH, 0.5, dpr);    // bottom edge
        ctx.moveTo(cx, by);
        ctx.lineTo(cx + cellW, by);
        if (ri === 0) {                            // top edge for first row
          const ty = cy + crispOffset(cy, 0.5, dpr);
          ctx.moveTo(cx, ty);
          ctx.lineTo(cx + cellW, ty);
        }
        if (ci === 0) {                            // left edge for first column
          const lx = cx + crispOffset(cx, 0.5, dpr);
          ctx.moveTo(lx, cy);
          ctx.lineTo(lx, cy + cellH);
        }
        ctx.stroke();
      }

      // Cell borders (base + any CF borders overlaid via per-edge merge).
      // For merged anchors, combine the anchor's border with the right/bottom
      // edges from the constituent cells on those edges (ECMA-376 §18.3.1.55).
      const baseBorder = mergeInfo
        ? resolveMergeBorder(border, rowIndex, colIndex, mergeInfo.right, mergeInfo.bottom, cellMap, styles)
        : border;
      let mergedBorder = mergeBorders(baseBorder, cf.border);
      // centerContinuous: hide internal vertical borders so the run reads as
      // one visual span (matches Excel — see precompute block above).
      if (suppressRightGridCol.has(ci) || suppressLeftGridCol.has(ci)) {
        mergedBorder = {
          ...mergedBorder,
          left: suppressLeftGridCol.has(ci) ? null : mergedBorder.left,
          right: suppressRightGridCol.has(ci) ? null : mergedBorder.right,
        };
      }
      // Shared-edge conflict resolution (ECMA-376 §18.18.3). Two adjacent cells
      // share the boundary along their common row / column edge. In the two-pass
      // order each cell's border closure runs after every fill, so the upper
      // cell's `bottom` (and the left cell's `right`) is never erased — it is
      // always drawn by that neighbour. So for an INTERIOR cell we only touch our
      // own top / left when *we* also define an edge there that conflicts, and
      // render the stronger of the two. Adopting a neighbour's edge we don't
      // define would double-stroke it (two strokes compounding to an over-dark
      // rule — sample-27) or paint a spurious line over a fill-less cell
      // (sample-6). This replaces the old fill-repair (`paintedFill` gate +
      // `invertedTop` / `invertedLeft`), which is gone.
      //
      // EXCEPTION — the quadrant's first rendered row/column (`ri`/`ci` === 0):
      // the upper / left neighbour lies OUTSIDE this quadrant (the viewport has
      // no top/left overscan — viewer.ts walks forward from startRow/startCol
      // with a +2 buffer only at the bottom/right), so when scrolled it is never
      // iterated and never strokes its own facing edge. There we adopt the
      // neighbour's edge even when our own side is unset, so a boundary authored
      // only as the neighbour's bottom/right still shows at the viewport edge.
      const aboveCell = cellMap.get(`${rowIndex - 1}:${colIndex}`);
      const aboveBottom = aboveCell
        ? resolveXf(styles, aboveCell.styleIndex ?? 0).border.bottom
        : null;
      if (aboveBottom?.style && (ri === 0 || mergedBorder.top?.style)) {
        mergedBorder = { ...mergedBorder, top: pickStrongerEdge(mergedBorder.top, aboveBottom) };
      }
      if (!suppressLeftGridCol.has(ci)) {
        // Skip when the left edge was deliberately suppressed for a
        // centerContinuous run (ECMA-376 §18.18.40) — otherwise the
        // neighbour's xf.right re-introduces the internal vertical we just hid.
        const leftCell = cellMap.get(`${rowIndex}:${colIndex - 1}`);
        const leftRight = leftCell
          ? resolveXf(styles, leftCell.styleIndex ?? 0).border.right
          : null;
        if (leftRight?.style && (ci === 0 || mergedBorder.left?.style)) {
          mergedBorder = { ...mergedBorder, left: pickStrongerEdge(mergedBorder.left, leftRight) };
        }
      }
      // Excel Table style overlay (ECMA-376 §18.8.83). Custom styles draw only
      // their dxf-defined borders; built-in style names fall back to a
      // synthesized accent rule. Drawn on top of the cell border so an
      // empty-border data cell still shows the table structure the style
      // defines. None-style tables produce no `tableStyleMap` entry, so this is
      // skipped. (Excel forbids merging cells inside a structured Table, so a
      // merged anchor never carries a `tableStyle`; for the merged branch only
      // the AutoFilter arrow can apply.)
      const overlay = tableStyle
        ? tableOverlayBorder(tableStyle, tsDxfWhole, tsDxfHeader, colIndex)
        : null;
      const drawArrow = rc.autoFilterCells.has(key);
      // Paint the base cell border, then the table overlay on top of it, then
      // the AutoFilter arrow — the original inline order, kept intact whether
      // this runs deferred (non-merged) or inline (merged anchor).
      const paintOverlayAndArrow = () => {
        if (overlay) {
          if (overlay.kind === 'dxf') {
            renderBorder(ctx, overlay.border, cx, cy, cellW, cellH, dpr);
          } else if (overlay.kind === 'accent') {
            const hp = 0.5 / dpr;
            ctx.strokeStyle = overlay.color;
            ctx.lineWidth = overlay.lineWidth;
            ctx.beginPath();
            // perimeter inset (-hp) kept literal: snapping could push it
            // outside the cell's bottom edge; dpr>=3 edge accepted
            ctx.moveTo(cx, cy + cellH - hp);
            ctx.lineTo(cx + cellW, cy + cellH - hp);
            if (overlay.topEdge) {
              const ty = cy + crispOffset(cy, overlay.lineWidth, dpr);
              ctx.moveTo(cx, ty);
              ctx.lineTo(cx + cellW, ty);
            }
            ctx.stroke();
          }
        }
        // AutoFilter dropdown indicator
        if (drawArrow) {
          drawAutoFilterArrow(ctx, cx, cy, cw, cellH);
        }
      };

      if (mergeInfo) {
        // Merged anchors draw their (full-span) border above all gridlines —
        // see `mergeBorderTasks`. Snapshot the per-iteration `let` binding
        // (`mergedBorder` is reassigned by the conflict-resolution above) so
        // the deferred closure draws the value as of *this* anchor; cx/cy/cellW/
        // cellH are already block `const` and safe to capture directly. The
        // overlay/arrow paint inline (pass 1) here, matching the prior order
        // where the deferred merge border sat above them.
        const mb = mergedBorder;
        mergeBorderTasks.push(() => renderBorder(ctx, mb, cx, cy, cellW, cellH, dpr));
        paintOverlayAndArrow();
      } else {
        // Non-merged borders are deferred to pass 2 (after every fill) so a
        // later neighbour's fill can never over-paint this cell's edges. The
        // overlay + arrow ride in the SAME closure, after the base border, so
        // their mutual base-border < overlay < arrow order is preserved exactly
        // — only the whole trio shifts above every cell's gray gridline.
        const mb = mergedBorder;
        borderTasks.push(() => {
          renderBorder(ctx, mb, cx, cy, cellW, cellH, dpr);
          paintOverlayAndArrow();
        });
      }

      if (!cell) continue;
      const text = formatCellValue(cell, styles, cf.numFmt, rc.worksheet.date1904);
      if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

      textTasks.push(() => {
      // The table-element dxf that applies to this cell's font (header/total/
      // column/stripe/wholeTable), resolved by the same §18.8.83 hierarchy as
      // the fill above. Used for both bold and color of custom styles.
      const tableFontDxf =
        (tableStyle?.isHeader) ? tsDxfHeader :
        (tableStyle?.isTotals) ? tsDxfTotal :
        (tableStyle?.isLastCol && tsDxfLastCol) ? tsDxfLastCol :
        (tableStyle?.isFirstCol && tsDxfFirstCol) ? tsDxfFirstCol :
        (tsDxfStripe) ? tsDxfStripe :
        (tableStyle ? tsDxfWhole : undefined);
      // Built-in style names: synthesize bold header/total rows (approximation).
      // Custom styles: bold comes strictly from the element dxf font (§18.5.1.2).
      const tableBold = tableStyle
        ? (tableStyle.isCustom
            ? !!tableFontDxf?.font?.bold
            : (tableStyle.isHeader || tableStyle.isTotals))
        : false;
      const effectiveBold = font.bold || !!cf.fontBold || tableBold;
      const effectiveItalic = font.italic || !!cf.fontItalic;
      const effectiveUnderline = font.underline || !!cf.fontUnderline;
      const effectiveStrike = font.strike || !!cf.fontStrike;
      const fontForDraw: CellFont = (
        effectiveBold !== font.bold || effectiveItalic !== font.italic ||
        effectiveUnderline !== font.underline || effectiveStrike !== font.strike
      )
        ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
        : font;
      ctx.font = buildFont(fontForDraw, cs);
      const hyperlinkUrl = rc.hyperlinkMap.get(key);
      // Table-style element dxfs can override font color (ECMA-376 §18.8.83),
      // following the same element hierarchy as the fill/bold above.
      const tableFontColor = tableFontDxf?.font?.color ?? null;
      const textColor = hyperlinkUrl
        ? '#0563C1'
        : (cf.fontColor ?? tableFontColor ?? font.color);
      ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const isNumeric = cell.value.type === 'number';
      const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';
      // Indent: ECMA-376 §18.8.1 alignment@indent — one level indents by 3
      // character widths (MDW) of the workbook's normal-style font.
      const indentPx = xf.indent ? Math.round(xf.indent * 3 * rc.mdw) : 0;
      // IconSet: reserve space on the left for the icon
      const iconSz = cf.iconSet ? Math.max(8, Math.round(Math.min(cellW, cellH) * 0.55)) : 0;
      const iconPad = iconSz > 0 ? iconSz + 4 : 0;
      const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;

      // ECMA-376 §18.18.40 ST_HorizontalAlignment value `centerContinuous`:
      // text is centered across the **selection range** — the leftmost cell
      // with content + adjacent empty cells to the right that also carry
      // `centerContinuous`. Cells stay independent (unlike merge), but the
      // centering uses the combined width. Walk right collecting empty
      // centerContinuous neighbours so we know how wide to center over.
      let centerContinuousW = cellW;
      let centerContinuousX = cx;
      let centerContinuousLastCi = ci;
      if (alignH === 'centerContinuous' && !mergeInfo) {
        for (let oci = ci + 1; oci < numCols; oci++) {
          const adjKey = `${rowIndex}:${startCol + oci}`;
          if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
          const adjCell = cellMap.get(adjKey);
          if (adjCell && adjCell.value.type !== 'empty') break;
          const adjStyleIndex = adjCell?.styleIndex ?? 0;
          const adjXf = resolveXf(styles, adjStyleIndex).xf;
          if (adjXf.alignH !== 'centerContinuous') break;
          centerContinuousW += colWidths[oci];
          centerContinuousLastCi = oci;
        }
      }

      // Text overflow into adjacent empty cells (ECMA-376 §18.3.1.4 "spans"
      // — Excel behavior when `wrapText=false`). Left-aligned text flows
      // rightward, right-aligned flows leftward, and centered splits evenly.
      // We only extend the clip rect; the text itself is still drawn once.
      // Stops at merge-cell boundaries, non-empty cells, and iconSet-left
      // overrun (since an icon sits inside this cell's left padding).
      let drawX = alignH === 'centerContinuous' ? centerContinuousX : cx;
      let drawW = alignH === 'centerContinuous' ? centerContinuousW : cellW;
      // Excel only overflows text into adjacent empty cells; numeric values
      // that don't fit are rendered as "####" (they never spill). Cells
      // containing hard line breaks render multi-line in place, so they don't
      // overflow either. centerContinuous overflows symmetrically across the
      // selection range — text stays centered on the original range, but the
      // clip rect extends into adjacent empty cells when the text is wider.
      const hasHardBreak = text.includes('\n');
      if (!mergeInfo && !xf.wrapText && !xf.textRotation && !isNumeric && !hasHardBreak) {
        const textW = ctx.measureText(text).width;
        const isCenterCont = alignH === 'centerContinuous';
        const textPx = isCenterCont ? textW + paddingX * 2 : textW + leftPad + paddingX;
        const containerW = isCenterCont ? centerContinuousW : cellW;
        if (textPx > containerW) {
          const overflow = textPx - containerW;
          let extendRight = 0;
          let extendLeft = 0;
          if (alignH === 'right') {
            extendLeft = overflow;
          } else if (alignH === 'center' || isCenterCont) {
            extendLeft = overflow / 2;
            extendRight = overflow / 2;
          } else {
            extendRight = overflow;
          }
          if (extendRight > 0) {
            let budget = extendRight;
            const startOci = isCenterCont ? centerContinuousLastCi + 1 : ci + 1;
            for (let oci = startOci; oci < numCols && budget > 0; oci++) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
          if (extendLeft > 0) {
            let budget = extendLeft;
            for (let oci = ci - 1; oci >= 0 && budget > 0; oci--) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawX -= colWidths[oci];
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
        }
      }

      // ECMA-376 §18.18.40 horizontal alignment:
      // - `fill`: repeat the cell content until the cell fills horizontally.
      //   We resolve to the repeated string here; layout is then standard
      //   left-aligned.
      // - `distributed`: distribute characters evenly across the cell width
      //   so the first char hugs the left edge and the last hugs the right.
      //   Implemented via Canvas `letterSpacing` (a string CSS length).
      // - `justify`: full-line justification. Multi-line distribution would
      //   need the wrap engine; for the single-line case we fall back to
      //   `distributed` if the text is short, else left-aligned (matches
      //   what Excel does when justify cannot expand a one-line string).
      // Anything we don't recognise falls through to left-align (general).
      let drawText = text;
      let letterSpacingPx = 0;
      if (alignH === 'fill' && !isNumeric && text.length > 0) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const oneW = ctx.measureText(text).width;
        if (oneW > 0 && oneW < innerW) {
          const reps = Math.max(1, Math.floor(innerW / oneW));
          drawText = text.repeat(reps);
        }
      }
      if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        const innerW = Math.max(1, cellW - paddingX * 2);
        const naturalW = ctx.measureText(drawText).width;
        const gaps = Math.max(1, [...drawText].length - 1);
        if (naturalW < innerW) {
          letterSpacingPx = Math.max(0, (innerW - naturalW) / gaps);
        }
      }

      let textX: number;
      let textAlign: CanvasTextAlign;
      if (alignH === 'right') {
        textX = cx + cellW - paddingX;
        textAlign = 'right';
      } else if (alignH === 'center') {
        textX = cx + cellW / 2;
        textAlign = 'center';
      } else if (alignH === 'centerContinuous') {
        // Center on the original selection range; drawX/drawW already covers
        // the range (and any overflow extension into empty neighbours).
        textX = centerContinuousX + centerContinuousW / 2;
        textAlign = 'center';
      } else if (alignH === 'distributed' || (alignH === 'justify' && !xf.wrapText && !hasHardBreak)) {
        // Distribute characters evenly: first hugs left edge, last hugs right.
        textX = cx + paddingX;
        textAlign = 'left';
      } else {
        textX = cx + leftPad;
        textAlign = 'left';
      }

      const rotation = xf.textRotation ?? 0;
      const isStacked = rotation === 255;
      const isRotated = rotation > 0 && rotation !== 255;

      // Draw icon set icon (before text clip block)
      if (cf.iconSet && iconSz > 0) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cx, cy, cellW, cellH);
        ctx.clip();
        drawCfIcon(ctx, cf.iconSet.name, cf.iconSet.index, cx + 2, cy + (cellH - iconSz) / 2, iconSz);
        ctx.restore();
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(drawX, cy, drawW, cellH);
      ctx.clip();

      // Stacked text (textRotation=255): draw each character on its own line
      if (isStacked) {
        const charH = vMetricPx(font.size, cs, 1.1);
        const totalH = text.length * charH;
        let charY = alignV === 'top' ? cy + paddingY
          : alignV === 'center' ? cy + (cellH - totalH) / 2
          : cy + cellH - totalH - paddingY;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        for (const ch of text) {
          ctx.fillText(ch, cx + cellW / 2, charY);
          charY += charH;
        }
        ctx.restore();
        return;
      }

      // Rotated text: translate to cell center, rotate, draw, restore
      if (isRotated) {
        const angleRad = rotation <= 90
          ? -(rotation * Math.PI / 180)
          : ((rotation - 90) * Math.PI / 180);
        ctx.translate(cx + cellW / 2, cy + cellH / 2);
        ctx.rotate(angleRad);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(text, 0, 0);
        ctx.restore();
        return;
      }

      // shrinkToFit: scale context horizontally if text is wider than cell
      if (xf.shrinkToFit) {
        const textW = ctx.measureText(text).width;
        const availW = cellW - leftPad - paddingX;
        if (textW > availW && textW > 0) {
          const scale = availW / textW;
          const pivotX = alignH === 'right' ? cx + cellW - paddingX
            : alignH === 'center' ? cx + cellW / 2
            : cx + leftPad;
          ctx.transform(scale, 0, 0, 1, pivotX * (1 - scale), 0);
        }
      }

      ctx.textAlign = textAlign;
      // ECMA-376 §18.18.39 distributed / §18.18.40 readingOrder. Canvas
      // `letterSpacing` (CSS length string, supported in modern browsers)
      // implements distributed by stretching the gap between glyphs;
      // `direction` flips the writing direction for RTL languages.
      if (letterSpacingPx > 0) {
        try { (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing = `${letterSpacingPx}px`; } catch { /* ignore */ }
      }
      // §18.8.1 readingOrder: 1 = LTR, 2 = RTL, 0/absent = Context (first
      // strong character decides). Always set the direction explicitly — the
      // previous explicit-only ladder leaked the prior cell's direction into
      // Context cells.
      try {
        (ctx as CanvasRenderingContext2D & { direction: 'ltr' | 'rtl' }).direction =
          cellBaseRtl(xf.readingOrder, text) ? 'rtl' : 'ltr';
      } catch { /* ignore */ }

      // Rich text: draw each run with its own font. Only supported for the
      // non-wrap path (wrap with mixed fonts is significantly more complex).
      const runs = cell.value.type === 'text' ? cell.value.runs : undefined;
      const hasRichText = runs && runs.length > 0;

      if (xf.wrapText && hasRichText) {
        // Rich text with wrapping — shared with the off-screen pre-pass.
        drawWrappedRichText(
          ctx, runs, fontForDraw,
          { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY },
          cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder },
        );
      } else if (xf.wrapText) {
        const lines = wrapTextLines(ctx, text, cellW - leftPad - paddingX);
        const lineH = vMetricPx(font.size, cs, 1.2, font.name ?? undefined);
        const totalTextH = lines.length * lineH;
        let startY: number;
        if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
        else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
        else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
        for (let li = 0; li < lines.length; li++) {
          ctx.fillText(lines[li], textX, startY + li * lineH);
        }
      } else if (hasRichText) {
        // Non-wrap rich text: per-run fonts, honoring hard breaks (Alt+Enter LF;
        // ECMA-376 §18.8.1 — Excel renders breaks even with wrapText off). The
        // shared helper draws a break-free value as a single alignV-anchored line
        // and a value with breaks as multiple lines, keeping this in-viewport
        // path and the off-screen-anchor pre-pass identical.
        drawNonWrapRichText(
          ctx, runs, fontForDraw,
          { alignH, alignV, cx, cy, cellW, cellH, leftPad, paddingX, paddingY },
          cs, dpr, { fontColor: cf.fontColor, readingOrder: xf.readingOrder },
        );
      } else {
        // ECMA-376 §18.4.14 vertAlign / ST_VerticalAlignRun §22.9.2.17 —
        // cell-level super/subscript: render the glyphs at ~65% size, shifted
        // off the baseline so the cell still reads at the right vertical band.
        // §18.4.14 mandates the size reduction; the exact ratio/offset is
        // implementation-defined (ratios match Office's visual output).
        const cellVertAlign = fontForDraw.vertAlign;
        const baseSizePxOrig = vMetricPx(font.size, cs);
        let vaYShift = 0;
        if (cellVertAlign === 'superscript') vaYShift = -Math.round(baseSizePxOrig * 0.35);
        else if (cellVertAlign === 'subscript') vaYShift = Math.round(baseSizePxOrig * 0.10);
        const drawFont: CellFont = cellVertAlign
          ? { ...fontForDraw, size: fontForDraw.size * 0.65 }
          : fontForDraw;
        if (cellVertAlign) {
          ctx.font = buildFont(drawFont, cs);
        }

        // Measure once for both underline and strike
        let overlayMetrics: TextMetrics | null = null;
        const measureOverlay = () => overlayMetrics ??= ctx.measureText(text);
        const overlayX = () => {
          const tW = Math.min(measureOverlay().width, drawW - leftPad - paddingX);
          return {
            x: alignH === 'right' ? cx + cellW - paddingX - tW
              : alignH === 'center' ? cx + cellW / 2 - tW / 2
              : cx + leftPad,
            width: tW,
          };
        };
        const sizePx = vMetricPx(drawFont.size, cs);

        if (fontForDraw.underline || hyperlinkUrl) {
          const { x: ux, width: tW } = overlayX();
          const uy = (alignV === 'top'
            ? cy + paddingY + sizePx + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(sizePx * 0.55)
              : cy + cellH - paddingY + 1) + vaYShift;
          const stroke = hyperlinkUrl ? '#0563C1' : (textColor ? hexToRgba(textColor) : '#000000');
          const dbl = fontForDraw.underlineStyle === 'double' || fontForDraw.underlineStyle === 'doubleAccounting';
          drawTextDecoLine(ctx, ux, ux + tW, uy, stroke, dbl, dpr);
        }
        if (fontForDraw.strike) {
          const { x: sx, width: tW } = overlayX();
          // Strike line sits roughly at the x-height mid-line (~45% up from baseline)
          const syBase = (alignV === 'top'
            ? cy + paddingY + Math.round(sizePx * 0.5)
            : alignV === 'center'
              ? cy + cellH / 2
              : cy + cellH - paddingY - Math.round(sizePx * 0.35)) + vaYShift;
          const sy = syBase + crispOffset(syBase, 0.5, dpr);
          ctx.save();
          ctx.strokeStyle = textColor ? hexToRgba(textColor) : '#000000';
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(sx + tW, sy); ctx.stroke();
          ctx.restore();
        }
        // Hard line breaks (\n from Alt+Enter) render as multiple lines even
        // when wrapText is false — this matches Excel's behavior.
        if (text.includes('\n')) {
          const lines = text.split('\n');
          const lineH = vMetricPx(font.size, cs, 1.2, font.name ?? undefined);
          const totalTextH = lines.length * lineH;
          let startY: number;
          if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
          else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
          else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
          for (let li = 0; li < lines.length; li++) {
            ctx.fillText(lines[li], textX, startY + li * lineH + vaYShift);
          }
        } else {
          let textY: number;
          if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
          else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
          else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
          ctx.fillText(drawText, textX, textY + vaYShift);
        }
      }

      ctx.restore();

      if (text && rc.onTextRun) {
        rc.onTextRun({ text, x: cx, y: cy, width: cellW, height: cellH, row: rowIndex, col: colIndex });
      }
      });
    }
  }

  // Two-pass flush. Pass 1 (the grid loop above) laid down every cell fill and
  // the gray sheet gridlines; pass 2 paints all styled borders on top, so no
  // fill can ever over-paint a neighbour's shared edge (the bug this removes).
  // Order within pass 2: non-merged borders first (they were inline = earlier
  // in the old per-cell path), then merged-anchor borders (always deferred =
  // later), then text last (stays on top of every border). This preserves the
  // prior gridline < non-merged border < merged border < text z-order.
  for (const task of borderTasks) task();

  for (const task of mergeBorderTasks) task();

  for (const task of textTasks) task();

  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Main render function
// ────────────────────────────────────────────────────────────────
/** Viewport-independent lookups derived purely from a Worksheet. The workbook
 *  caches one Worksheet object per sheet, so this WeakMap hits on every scroll
 *  frame and only misses on a sheet switch / re-parse — avoiding a full-sheet
 *  cell-Map rebuild and a conditional-formatting recompile per frame. */
interface SheetRenderCache {
  cellMap: Map<string, Cell>;
  cfContext: CfContext;
  mergeSkipSet: Set<string>;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
  commentCells: Set<string>;
  tableStyleMap: Map<string, TableCellStyle>;
  sparklineMap: Map<string, SparklineModel>;
}
const sheetRenderCache = new WeakMap<Worksheet, SheetRenderCache>();

function getSheetRenderCache(worksheet: Worksheet): SheetRenderCache {
  const cached = sheetRenderCache.get(worksheet);
  if (cached) return cached;

  const cellMap = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      cellMap.set(`${cell.row}:${cell.col}`, cell);
    }
  }

  // Merge skip-set is pure topology (no scaled sizes) so it is cacheable; the
  // anchor map's pixel sizes depend on cellScale and stay per-frame.
  const mergeSkipSet = new Set<string>();
  for (const mc of worksheet.mergeCells ?? []) {
    for (let r = mc.top; r <= mc.bottom; r++) {
      for (let c = mc.left; c <= mc.right; c++) {
        if (r === mc.top && c === mc.left) continue;
        mergeSkipSet.add(`${r}:${c}`);
      }
    }
  }

  const autoFilterCells = new Set<string>();
  if (worksheet.autoFilter) {
    const af = worksheet.autoFilter;
    for (let c = af.left; c <= af.right; c++) autoFilterCells.add(`${af.top}:${c}`);
  }

  const hyperlinkMap = new Map<string, string>();
  for (const hl of worksheet.hyperlinks ?? []) {
    if (hl.url) hyperlinkMap.set(`${hl.row}:${hl.col}`, hl.url);
  }

  const commentCells = new Set<string>();
  for (const ref of worksheet.commentRefs ?? []) {
    const parsed = parseA1Ref(ref);
    if (parsed) commentCells.add(`${parsed.row}:${parsed.col}`);
  }

  const entry: SheetRenderCache = {
    cellMap,
    cfContext: compileCf(worksheet),
    mergeSkipSet,
    autoFilterCells,
    hyperlinkMap,
    commentCells,
    tableStyleMap: buildTableStyleMap(worksheet),
    sparklineMap: buildSparklineMap(worksheet),
  };
  sheetRenderCache.set(worksheet, entry);
  return entry;
}

export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): void {
  const dpr = opts.dpr ?? 1;
  const cs = opts.cellScale ?? 1;
  // Resolve MDW once per render — workbook-wide value derived from the
  // Normal-style font (ECMA-376 §18.3.1.13). Cached internally per
  // (family, sizePt) so repeated renders are O(1).
  const mdw = getMdwForWorksheet(worksheet);
  const canvasW = ctx.canvas.width / dpr;
  const canvasH = ctx.canvas.height / dpr;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  // Scaled pixel helper: apply cellScale to all cell/header dimensions
  const sp = (px: number) => Math.round(px * cs);
  const hw = sp(HEADER_W);  // scaled header column width
  const hh = sp(HEADER_H);  // scaled header row height

  const { row: startRow, col: startCol, rows: numRows, cols: numCols } = viewport;
  const scrollOffsetX = (opts.scrollOffsetX ?? 0) * cs;
  const scrollOffsetY = (opts.scrollOffsetY ?? 0) * cs;
  const freezeRows = opts.freezeRows ?? 0;
  const freezeCols = opts.freezeCols ?? 0;

  // ── Compute frozen area pixel sizes (scaled) ─────────────────
  const frozenColWidths: number[] = [];
  for (let c = 1; c <= freezeCols; c++) {
    frozenColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw)));
  }
  const frozenRowHeights: number[] = [];
  for (let r = 1; r <= freezeRows; r++) {
    frozenRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }
  const frozenW = frozenColWidths.reduce((s, w) => s + w, 0);
  const frozenH = frozenRowHeights.reduce((s, h) => s + h, 0);

  // ── Scrollable col/row pixel widths (scaled) ─────────────────
  const scrollColWidths: number[] = [];
  for (let c = startCol; c < startCol + numCols; c++) {
    scrollColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw)));
  }
  const scrollRowHeights: number[] = [];
  for (let r = startRow; r < startRow + numRows; r++) {
    scrollRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }

  // ── Viewport-independent lookups (memoized per Worksheet) ────
  const {
    cellMap, cfContext, mergeSkipSet, autoFilterCells,
    hyperlinkMap, commentCells, tableStyleMap, sparklineMap,
  } = getSheetRenderCache(worksheet);

  // Merge anchor sizes are cellScale-scaled, so they stay per-frame.
  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number; right: number; bottom: number }>();
  for (const mc of worksheet.mergeCells ?? []) {
    let totalW = 0;
    for (let c = mc.left; c <= mc.right; c++) {
      totalW += sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth, mdw));
    }
    let totalH = 0;
    for (let r = mc.top; r <= mc.bottom; r++) {
      totalH += sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight));
    }
    mergeAnchorMap.set(`${mc.top}:${mc.left}`, { totalW, totalH, right: mc.right, bottom: mc.bottom });
  }

  const rc: RenderContext = {
    worksheet, styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext,
    colWidths: scrollColWidths,
    rowHeights: scrollRowHeights,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    startRow, startCol,
    cs,
    dpr,
    autoFilterCells,
    hyperlinkMap,
    commentCells,
    tableStyleMap,
    sparklineMap,
    mdw,
    onTextRun: opts.onTextRun,
    rtl: worksheet.rightToLeft === true,
    canvasW,
  };

  // Canvas areas for each quadrant
  const cellAreaX = hw;
  const cellAreaY = hh;
  const scrollAreaX = cellAreaX + frozenW;
  const scrollAreaY = cellAreaY + frozenH;
  const scrollAreaW = Math.max(0, canvasW - scrollAreaX);
  const scrollAreaH = Math.max(0, canvasH - scrollAreaY);

  // ── Q1: frozen rows × frozen cols ───────────────────────────
  if (freezeRows > 0 && freezeCols > 0) {
    renderQuadrant(ctx, rc,
      1, 1, frozenColWidths, frozenRowHeights,
      0, 0,
      cellAreaX, cellAreaY,
      cellAreaX, cellAreaY, frozenW, frozenH,
    );
  }

  // ── Q2: frozen rows × scrollable cols ───────────────────────
  if (freezeRows > 0) {
    renderQuadrant(ctx, rc,
      1, startCol, scrollColWidths, frozenRowHeights,
      scrollOffsetX, 0,
      scrollAreaX, cellAreaY,
      scrollAreaX, cellAreaY, scrollAreaW, frozenH,
    );
  }

  // ── Q3: scrollable rows × frozen cols ───────────────────────
  if (freezeCols > 0) {
    renderQuadrant(ctx, rc,
      startRow, 1, frozenColWidths, scrollRowHeights,
      0, scrollOffsetY,
      cellAreaX, scrollAreaY,
      cellAreaX, scrollAreaY, frozenW, scrollAreaH,
    );
  }

  // ── Q4: scrollable rows × scrollable cols (main area) ───────
  renderQuadrant(ctx, rc,
    startRow, startCol, scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    scrollAreaX, scrollAreaY,
    scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH,
  );

  // ── Anchored images (clipped to scrollable area) ─────────────
  if (worksheet.images && worksheet.images.length > 0 && opts.loadedImages) {
    renderImages(
      ctx, worksheet, opts.loadedImages, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Anchored shape groups (custom geometry, incl. embedded images) ────
  if (worksheet.shapeGroups && worksheet.shapeGroups.length > 0) {
    renderShapeGroups(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
      opts.loadedImages,
    );
  }

  // ── Anchored charts (clipped to scrollable area) ──────────────
  if (worksheet.charts && worksheet.charts.length > 0) {
    renderCharts(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Anchored slicers (Office 2010+ pivot/table filter buttons) ──
  if (worksheet.slicers && worksheet.slicers.length > 0) {
    renderSlicers(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Row/col headers (drawn last, always on top) ──────────────
  renderHeaders(ctx, canvasW, canvasH,
    startRow, startCol, numRows, numCols,
    scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    hw, hh, cs, dpr,
    opts.selectedRowRange ?? null,
    opts.selectedColRange ?? null,
    worksheet.rightToLeft === true,
  );

  // ── Freeze pane separator lines ──────────────────────────────
  // RTL mirrors the vertical freeze divider to the left edge of the frozen
  // band (which now anchors at the right). The horizontal divider is
  // unaffected by the x-mirror but its row-header end moves to the right.
  const rtl = worksheet.rightToLeft === true;
  // crispOffset snaps the 0.5-px divider onto the device grid (same convention
  // as gridlines/headers). A fixed 0.5/dpr offset blurred at even device
  // widths (e.g. dpr=3).
  if (freezeRows > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const dy = scrollAreaY + crispOffset(scrollAreaY, 0.5, dpr);
    if (rtl) {
      // cell area is [0, canvasW - hw]
      ctx.moveTo(0, dy);
      ctx.lineTo(canvasW - hw, dy);
    } else {
      ctx.moveTo(hw, dy);
      ctx.lineTo(canvasW, dy);
    }
    ctx.stroke();
    ctx.restore();
  }
  if (freezeCols > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    // Divider sits between the frozen band and the scroll band. In LTR that
    // is at scrollAreaX = hw + frozenW; mirrored that becomes
    // canvasW - scrollAreaX.
    const dividerBase = rtl ? canvasW - scrollAreaX : scrollAreaX;
    const dividerX = dividerBase + crispOffset(dividerBase, 0.5, dpr);
    ctx.moveTo(dividerX, hh);
    ctx.lineTo(dividerX, canvasH);
    ctx.stroke();
    ctx.restore();
  }
}

// ────────────────────────────────────────────────────────────────
// Headers
// ────────────────────────────────────────────────────────────────
function renderHeaders(
  ctx: CanvasRenderingContext2D,
  canvasW: number, canvasH: number,
  startRow: number, startCol: number,
  numRows: number, numCols: number,
  scrollColWidths: number[], scrollRowHeights: number[],
  scrollOffsetX: number, scrollOffsetY: number,
  frozenColWidths: number[], frozenRowHeights: number[],
  frozenW: number, frozenH: number,
  hw: number, hh: number, cs: number, dpr: number,
  selectedRowRange: { start: number; end: number; strong: boolean } | null,
  selectedColRange: { start: number; end: number; strong: boolean } | null,
  rtl: boolean,
): void {
  const HEADER_BG = '#f8f9fa';
  const HEADER_BG_SUBTLE = '#e8eaed';
  const HEADER_BG_STRONG = '#caddf6';
  const HEADER_BORDER = '#c8ccd0';
  const HEADER_BORDER_STRONG = '#5b9bd5';
  const HEADER_TEXT = '#444';

  const colBg = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BG;
    return selectedColRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const colBorder = (col: number): string => {
    if (!selectedColRange || col < selectedColRange.start || col > selectedColRange.end) return HEADER_BORDER;
    return selectedColRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const rowBg = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BG;
    return selectedRowRange.strong ? HEADER_BG_STRONG : HEADER_BG_SUBTLE;
  };
  const rowBorder = (row: number): string => {
    if (!selectedRowRange || row < selectedRowRange.start || row > selectedRowRange.end) return HEADER_BORDER;
    return selectedRowRange.strong ? HEADER_BORDER_STRONG : HEADER_BORDER;
  };
  const headerFontSize = Math.max(1, Math.round(11 * cs));
  const HEADER_FONT = `${headerFontSize}px ${DEFAULT_FONT_FAMILY}`;
  const scrollAreaX = hw + frozenW;
  const scrollAreaY = hh + frozenH;
  const hp = 0.5 / dpr;  // half device-pixel offset for 1dp crisp lines

  // RTL grid mirror (ECMA-376 §18.3.1.87). The cell-area band mirrors about
  // canvasW: a left-anchored col-header rect [x, x+w] maps to its mirror.
  // The corner box and the row-number strip move from the left edge
  // ([0, hw]) to the right edge ([canvasW - hw, canvasW]).
  const mirrorX = (x: number, w: number): number => rtl ? rtlMirrorX(x, w, canvasW) : x;
  const cornerX = rtl ? canvasW - hw : 0;  // x-origin of the corner / row-header strip

  // Corner – draw all 4 edges (standalone box)
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(cornerX, 0, hw, hh);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  const cornerL = cornerX + crispOffset(cornerX, 0.5, dpr);
  const cornerT = crispOffset(0, 0.5, dpr);
  ctx.moveTo(cornerL, 0); ctx.lineTo(cornerL, hh);                    // left
  ctx.moveTo(cornerX, cornerT); ctx.lineTo(cornerX + hw, cornerT);    // top
  // perimeter inset (-hp) kept literal: snapping could push it outside the
  // header box; dpr>=3 edge accepted
  ctx.moveTo(cornerX + hw - hp, 0); ctx.lineTo(cornerX + hw - hp, hh);  // right
  ctx.moveTo(cornerX, hh - hp); ctx.lineTo(cornerX + hw, hh - hp);    // bottom
  ctx.stroke();

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;

  // Helper: draw one column header cell.
  // The inter-column separator is the LEADING (left) edge at +hp, so it lands on
  // the SAME device pixel as the data grid's vertical column boundary (gridlines
  // and cell borders draw the shared boundary at cx+hp). Drawing it as this
  // cell's own left edge — rather than the previous cell's trailing right edge at
  // cx+cw-hp — both aligns it with the data column line (the strip used to sit 1
  // device px to the left) AND survives the next cell's fillRect (which starts at
  // cx+cw and never touches cx+hp), so no separate fill/stroke pass is needed.
  // The top/bottom lines stay as the header-strip perimeter.
  const drawColHeader = (col: number, ltrCx: number, cw: number) => {
    const cx = mirrorX(ltrCx, cw);
    ctx.fillStyle = colBg(col);
    ctx.fillRect(cx, 0, cw, hh);
    ctx.strokeStyle = colBorder(col);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const colSepX = cx + crispOffset(cx, 0.5, dpr);  // column boundary, aligns with data grid
    const colTopY = crispOffset(0, 0.5, dpr);
    ctx.moveTo(colSepX, 0);           ctx.lineTo(colSepX, hh);       // left = column boundary (aligns with data grid)
    ctx.moveTo(cx, hh - hp);          ctx.lineTo(cx + cw, hh - hp);  // bottom (strip perimeter, inset -hp kept literal)
    ctx.moveTo(cx, colTopY);          ctx.lineTo(cx + cw, colTopY);  // top (strip perimeter)
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colToLetter(col), cx + cw / 2, hh / 2);
  };

  // Helper: draw one row header cell.
  // The inter-row separator is the LEADING (top) edge at +hp, so it lands on the
  // SAME device pixel as the data grid's horizontal row boundary (gridlines and
  // cell borders draw the shared boundary at cy+hp). Drawing it as this cell's
  // own top edge — rather than the previous cell's trailing bottom edge at
  // cy+ch-hp — aligns it with the data row line (the strip used to sit 1 device
  // px above) AND survives the next cell's fillRect (which starts at cy+ch and
  // never touches cy+hp), so no separate fill/stroke pass is needed. The
  // left/right lines stay as the header-strip perimeter.
  const drawRowHeader = (row: number, cy: number, ch: number) => {
    const rx = cornerX;  // left edge of the row-header strip (mirrors to the right in RTL)
    ctx.fillStyle = rowBg(row);
    ctx.fillRect(rx, cy, hw, ch);
    ctx.strokeStyle = rowBorder(row);
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    const rowSepY = cy + crispOffset(cy, 0.5, dpr);   // row boundary, aligns with data grid
    const rowLeftX = rx + crispOffset(rx, 0.5, dpr);
    ctx.moveTo(rx + hw - hp, cy);  ctx.lineTo(rx + hw - hp, cy + ch);   // right (strip perimeter, inset -hp kept literal)
    ctx.moveTo(rx, rowSepY);       ctx.lineTo(rx + hw, rowSepY);         // top = row boundary (aligns with data grid)
    ctx.moveTo(rowLeftX, cy);      ctx.lineTo(rowLeftX, cy + ch);        // left (strip perimeter)
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textBaseline = 'middle';
    const pad = Math.max(2, Math.round(4 * cs));
    if (rtl) {
      // Strip is on the right; the grid sits to its left, so anchor the
      // number toward the grid-facing (left) edge.
      ctx.textAlign = 'left';
      ctx.fillText(String(row), rx + pad, cy + ch / 2);
    } else {
      ctx.textAlign = 'right';
      ctx.fillText(String(row), rx + hw - pad, cy + ch / 2);
    }
  };

  // Frozen col headers (no h-scroll, fixed positions)
  if (frozenColWidths.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(mirrorX(hw, frozenW), 0, frozenW, hh);
    ctx.clip();
    let cx = hw;
    for (let ci = 0; ci < frozenColWidths.length; ci++) {
      drawColHeader(ci + 1, cx, frozenColWidths[ci]);
      cx += frozenColWidths[ci];
    }
    ctx.restore();
  }

  // Scrollable col headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(mirrorX(scrollAreaX, canvasW - scrollAreaX), 0, canvasW - scrollAreaX, hh);
  ctx.clip();
  let cx = scrollAreaX - scrollOffsetX;
  for (let ci = 0; ci < scrollColWidths.length; ci++) {
    const cw = scrollColWidths[ci];
    if (cx + cw > scrollAreaX && cx < canvasW) {
      drawColHeader(startCol + ci, cx, cw);
    }
    cx += cw;
  }
  ctx.restore();

  // Frozen row headers (no v-scroll)
  if (frozenRowHeights.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(cornerX, hh, hw, frozenH);
    ctx.clip();
    let cy = hh;
    for (let ri = 0; ri < frozenRowHeights.length; ri++) {
      drawRowHeader(ri + 1, cy, frozenRowHeights[ri]);
      cy += frozenRowHeights[ri];
    }
    ctx.restore();
  }

  // Scrollable row headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(cornerX, scrollAreaY, hw, canvasH - scrollAreaY);
  ctx.clip();
  let cy = scrollAreaY - scrollOffsetY;
  for (let ri = 0; ri < scrollRowHeights.length; ri++) {
    const ch = scrollRowHeights[ri];
    if (cy + ch > scrollAreaY && cy < canvasH) {
      drawRowHeader(startRow + ri, cy, ch);
    }
    cy += ch;
  }
  ctx.restore();

}

// ────────────────────────────────────────────────────────────────
// Image anchors  (ECMA-376 §20.5, <xdr:twoCellAnchor>)
// ────────────────────────────────────────────────────────────────

/** Sum scaled column widths for cols 1..n-1 (sheet-space X of col n in scaled px). */
function sheetXForCol(
  ws: Worksheet,
  col1: number, // 1-indexed column number
  cs: number,
): number {
  const mdw = getMdwForWorksheet(ws);
  let x = 0;
  for (let c = 1; c < col1; c++) {
    x += Math.round(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw) * cs);
  }
  return x;
}

/** Sum scaled row heights for rows 1..n-1 (sheet-space Y of row n in scaled px). */
function sheetYForRow(
  ws: Worksheet,
  row1: number, // 1-indexed row number
  cs: number,
): number {
  let y = 0;
  for (let r = 1; r < row1; r++) {
    y += Math.round(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight) * cs);
  }
  return y;
}


function renderImages(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  loadedImages: Map<string, CanvasImageSource | null>,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  // Sheet-space origin of the current scroll viewport's first visible cell
  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of ws.images) {
    const img = loadedImages.get(anchor.imagePath);
    if (!img) continue;

    // xdr col/row are 0-indexed; our widths map is 1-indexed.
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;

    // Image sheet-space top-left (always derived from the `from` anchor)
    const imgSheetX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const imgSheetY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;

    // ECMA-376 §20.5.2.33 + "Move but don't size with cells": when the
    // anchor was saved with editAs="oneCell" Excel preserves the picture's
    // saved EMU size (<xdr:spPr><a:xfrm><a:ext>) regardless of cell
    // resizing, and the to anchor is only updated to track that fixed size.
    // Use the native ext directly so the rendered image matches Excel even
    // when our column-width / row-height computation diverges slightly from
    // Excel's (e.g. row ht is stored as px in this viewer but Excel applies
    // pt→px for some files). Falls back to the from/to-derived rect for
    // editAs="twoCell" (default, image resizes with cells) and absolute
    // anchors, or when the parser couldn't capture the native ext.
    let imgW: number, imgH: number;
    if (anchor.editAs === 'oneCell' && anchor.nativeExtCx > 0 && anchor.nativeExtCy > 0) {
      imgW = (anchor.nativeExtCx * cs) / EMU_PER_PX;
      imgH = (anchor.nativeExtCy * cs) / EMU_PER_PX;
    } else {
      const toCol1 = anchor.toCol + 1;
      const toRow1 = anchor.toRow + 1;
      const imgSheetX2 = sheetXForCol(ws, toCol1, cs) + (anchor.toColOff * cs) / EMU_PER_PX;
      const imgSheetY2 = sheetYForRow(ws, toRow1, cs) + (anchor.toRowOff * cs) / EMU_PER_PX;
      imgW = imgSheetX2 - imgSheetX1;
      imgH = imgSheetY2 - imgSheetY1;
    }
    if (imgW <= 0 || imgH <= 0) continue;

    // Translate to canvas coordinates of the scrollable viewport
    const canvasX = scrollAreaX + (imgSheetX1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (imgSheetY1 - scrollOriginSheetY) - scrollOffsetY;

    // Early out when entirely off-screen
    if (canvasX + imgW < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + imgH < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    drawImageCropped(ctx, img, anchor.srcRect, canvasX, canvasY, imgW, imgH);
  }

  ctx.restore();
}

function renderShapeGroups(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
  loadedImages?: Map<string, CanvasImageSource | null>,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const anchors = ws.shapeGroups;
  if (!anchors || anchors.length === 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of anchors) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;

    const x1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const y1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;

    // editAs="oneCell" preserves the group's saved grpSpPr/xfrm/ext EMU
    // size regardless of cell resizing (ECMA-376 §20.5.2.33). See
    // renderImages for the same handling on stand-alone <xdr:pic>.
    let w: number, h: number;
    if (anchor.editAs === 'oneCell' && anchor.nativeExtCx > 0 && anchor.nativeExtCy > 0) {
      w = (anchor.nativeExtCx * cs) / EMU_PER_PX;
      h = (anchor.nativeExtCy * cs) / EMU_PER_PX;
    } else {
      const toCol1 = anchor.toCol + 1;
      const toRow1 = anchor.toRow + 1;
      const x2 = sheetXForCol(ws, toCol1, cs) + (anchor.toColOff * cs) / EMU_PER_PX;
      const y2 = sheetYForRow(ws, toRow1, cs) + (anchor.toRowOff * cs) / EMU_PER_PX;
      w = x2 - x1;
      h = y2 - y1;
    }
    if (w <= 0 || h <= 0) continue;

    const canvasX = scrollAreaX + (x1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (y1 - scrollOriginSheetY) - scrollOffsetY;

    if (canvasX + w < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + h < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    for (const shape of anchor.shapes) {
      const sx = canvasX + shape.x * w;
      const sy = canvasY + shape.y * h;
      const sw = shape.w * w;
      const sh = shape.h * h;
      if (sw <= 0 || sh <= 0) continue;
      drawShape(ctx, shape, sx, sy, sw, sh, cs, loadedImages);
    }
  }

  ctx.restore();
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  sx: number, sy: number, sw: number, sh: number,
  cs: number,
  loadedImages?: Map<string, CanvasImageSource | null>,
): void {
  ctx.save();
  if (shape.rot !== 0) {
    ctx.translate(sx + sw / 2, sy + sh / 2);
    ctx.rotate((shape.rot * Math.PI) / 180);
    ctx.translate(-sw / 2, -sh / 2);
  } else {
    ctx.translate(sx, sy);
  }

  if (shape.geom.type === 'custom') {
    for (const path of shape.geom.paths) {
      if (path.w <= 0 || path.h <= 0) continue;
      const kx = sw / path.w;
      const ky = sh / path.h;
      ctx.beginPath();
      // Track pen position for arcTo center computation.
      let penX = 0, penY = 0;
      // Track subpath start for close lineTo.
      let subX = 0, subY = 0;
      for (const cmd of path.commands) {
        switch (cmd.op) {
          case 'moveTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.moveTo(px, py);
            penX = subX = px; penY = subY = py;
            break;
          }
          case 'lineTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.lineTo(px, py);
            penX = px; penY = py;
            break;
          }
          case 'cubicBezTo': {
            const ex = cmd.x3 * kx, ey = cmd.y3 * ky;
            ctx.bezierCurveTo(
              cmd.x1 * kx, cmd.y1 * ky,
              cmd.x2 * kx, cmd.y2 * ky,
              ex, ey,
            );
            penX = ex; penY = ey;
            break;
          }
          case 'quadBezTo': {
            const ex = cmd.x2 * kx, ey = cmd.y2 * ky;
            ctx.quadraticCurveTo(cmd.x1 * kx, cmd.y1 * ky, ex, ey);
            penX = ex; penY = ey;
            break;
          }
          case 'arcTo': {
            // ECMA-376 §20.1.9.3: pen lies on ellipse at stAng;
            // derive center from pen + stAng, then sweep swAng.
            const rx = cmd.wr * kx, ry = cmd.hr * ky;
            if (rx <= 0 || ry <= 0) break;
            const stRad = (cmd.stAng / 60000) * (Math.PI / 180);
            const swRad = (cmd.swAng / 60000) * (Math.PI / 180);
            const cx = penX - Math.cos(stRad) * rx;
            const cy = penY - Math.sin(stRad) * ry;
            const endRad = stRad + swRad;
            ctx.ellipse(cx, cy, rx, ry, 0, stRad, endRad, swRad < 0);
            penX = cx + Math.cos(endRad) * rx;
            penY = cy + Math.sin(endRad) * ry;
            break;
          }
          case 'close':
            ctx.closePath();
            penX = subX; penY = subY;
            break;
        }
      }
      fillAndStroke(ctx, shape);
    }
  } else if (shape.geom.type === 'preset') {
    // Drive the shape off the ECMA-376 §20.1.9 spec-driven preset engine
    // (presets.json from presetShapeDefinitions.xml). It honours each path's
    // own fill/stroke attributes, so a parallelogram / rtTriangle / callout
    // renders with its true outline instead of the old rect fallback. The
    // engine returns false for presets it doesn't carry; only then do we fall
    // back to a plain rectangle.
    const baseFill = shape.fillColor ?? null;
    const applyAndStroke =
      shape.strokeColor && shape.strokeWidth > 0
        ? () => {
            ctx.strokeStyle = shape.strokeColor as string;
            ctx.lineWidth = Math.max(0.5, shape.strokeWidth / EMU_PER_PX);
            ctx.stroke();
          }
        : null;
    const drawn = renderPresetShape(
      ctx,
      shape.geom.name,
      0,
      0,
      sw,
      sh,
      shape.geom.adj ?? [],
      baseFill,
      applyAndStroke,
      // xlsx shapes carry no drop shadow on the body, so there's nothing to
      // clear between the engine's fill and stroke passes.
      () => {},
    );
    if (!drawn) {
      ctx.beginPath();
      ctx.rect(0, 0, sw, sh);
      fillAndStroke(ctx, shape);
    }
  } else if (shape.geom.type === 'image') {
    // Image leaf inside a group (e.g. a sun-emoji clip-art nested in the
    // calendar header). The caller pre-decodes every image path seen in
    // `ws.shapeGroups[*].shapes[*].geom` via XlsxWorkbook.renderViewport,
    // so we should normally have it in `loadedImages` (keyed by imagePath).
    // If not, fall back to a silent skip — drawing an empty rect would look
    // worse.
    const img = loadedImages?.get(shape.geom.imagePath);
    if (img) {
      // Honor an `<a:srcRect>` crop on the leaf pic (oneCellAnchor / grpSp leaf),
      // same as the top-level anchor path (ECMA-376 §20.1.8.55).
      drawImageCropped(ctx, img, shape.geom.srcRect, 0, 0, sw, sh);
    }
  }
  // Shape text body (ECMA-376 §20.5.2.34 `<xdr:txBody>`). Drawn after
  // fill/stroke so it sits on top of the shape's background.
  if (shape.text) {
    drawShapeText(ctx, shape.text, sw, sh, cs);
  }
  ctx.restore();
}

/**
 * Render a shape's `<xdr:txBody>` content into the local (already-translated)
 * coordinate system spanning [0..sw, 0..sh].
 *
 * Handles only the subset that real-world Excel files exercise heavily —
 * single column of paragraphs, per-run bold/italic/size/color/font, paragraph
 * align (`l`/`ctr`/`r`), body anchor (`t`/`ctr`/`b`). Text wrapping uses
 * canvas measurements when `bodyPr@wrap="square"` (the default), and the
 * inset is the OOXML default (`lIns=91440 EMU` ≈ 7.2 pt on each side, plus
 * `tIns=45720 EMU` ≈ 3.6 pt top/bottom). We approximate inset as a fixed
 * 7 px / 4 px since `bodyPr@*Ins` is rarely overridden in practice.
 */
// ── Math (OMML) rendering in shapes ─────────────────────────────────────────
// Equations are converted to SVG by MathJax once, cached by their MathNode[]
// reference (stable from parse through render), then drawn as images inside
// shape text. Mirrors the pptx/docx pipeline so typesetting is identical; the
// engine is opt-in (injected once at XlsxWorkbook.load via LoadOptions.math,
// stored on the instance) and tree-shakes out otherwise.
interface MathRender {
  /** The equation rasterized as opaque black glyphs on transparent. */
  img: HTMLImageElement;
  /** baseline-relative extents in em (1em = the equation's font size in px). */
  widthEm: number;
  ascentEm: number;
  descentEm: number;
  /** Per-colour tinted copies (the black `img` recoloured via source-in). */
  tinted: Map<string, CanvasImageSource | null>;
}
const mathRenders = new WeakMap<MathNode[], MathRender>();

/** Tint the cached black equation image to `color` via source-in (cached). */
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

// px-per-em rasterization resolution for equation SVGs — keeps glyphs crisp on
// HiDPI canvases (a MathJax SVG otherwise rasterizes at a small intrinsic size
// and drawImage upscales it). 256 stays crisp past 40pt at devicePixelRatio 3.
const MATH_RASTER_PX_PER_EM = 256;
function sizeSvgForRaster(svg: string, widthEm: number, heightEm: number): string {
  const w = Math.max(1, Math.round(widthEm * MATH_RASTER_PX_PER_EM));
  const h = Math.max(1, Math.round(heightEm * MATH_RASTER_PX_PER_EM));
  return svg.replace(/<svg([^>]*?)>/, (_m, attrs: string) => {
    const cleaned = attrs.replace(/\s(?:width|height)="[^"]*"/g, '');
    return `<svg${cleaned} width="${w}" height="${h}">`;
  });
}

/** Gather every math run reachable from a worksheet's shapes. Equations live
 *  only in shape text bodies (ECMA-376 §22.1) — never in cell values. */
function collectWorksheetMath(ws: Worksheet): { nodes: MathNode[]; display: boolean }[] {
  const found: { nodes: MathNode[]; display: boolean }[] = [];
  for (const anchor of ws.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      for (const p of shape.text?.paragraphs ?? []) {
        for (const run of p.runs) {
          if (run.type === 'math') found.push({ nodes: run.nodes, display: run.display });
        }
      }
    }
  }
  return found;
}

/** True iff the worksheet has at least one equation not yet rasterized. Lets
 *  the caller keep steady-state scroll/zoom frames fully synchronous (no await)
 *  — only the first frame that reveals new equations pays the async cost. */
export function worksheetHasUncachedMath(ws: Worksheet): boolean {
  for (const anchor of ws.shapeGroups ?? []) {
    for (const shape of anchor.shapes) {
      for (const p of shape.text?.paragraphs ?? []) {
        for (const run of p.runs) {
          if (run.type === 'math' && !mathRenders.has(run.nodes)) return true;
        }
      }
    }
  }
  return false;
}

/** Pre-rasterize every equation in the worksheet (async). Idempotent: skips
 *  already-rasterized equations and only loads MathJax when math is present.
 *  MUST be awaited BEFORE the canvas resize (off the synchronous draw path), so
 *  the old frame stays visible until the new one is ready (no white flash). */
export async function prepareWorksheetMath(ws: Worksheet, math: MathRenderer): Promise<void> {
  const uncached = collectWorksheetMath(ws).filter((r) => !mathRenders.has(r.nodes));
  if (uncached.length === 0) return;
  await math.loadMathJax();
  for (const r of uncached) {
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
      // Conversion failure: leave the equation unrendered rather than throw.
    }
  }
}

export function drawShapeText(
  ctx: CanvasRenderingContext2D,
  txt: import('./types.js').ShapeText,
  sw: number, sh: number,
  cs: number,
): void {
  if (sw <= 0 || sh <= 0 || txt.paragraphs.length === 0) return;
  // The shape box (sw,sh) is already cellScale-scaled by the caller, but font
  // sizes are authored in points. Multiply every px size (padding, text, math)
  // by `cs` so the contents grow/shrink with the box on Excel's zoom slider —
  // matching how cell text scales via buildFont(font, cs). Without this the box
  // would zoom while the glyphs stayed fixed, drifting out of alignment.
  // Text insets from `<a:bodyPr>` (ECMA-376 §21.1.2.1.1), EMU → px, scaled by
  // `cs` like every other px size here. The parser always emits them (spec
  // default 91440 EMU L/R = 9.6 px, 45720 EMU T/B = 4.8 px). This replaces the
  // former empirical padX=7 / padY=4 constants with the real, per-side insets —
  // the box is inset asymmetrically when the shape authors distinct lIns/rIns or
  // tIns/bIns.
  const padLeft = (txt.lIns / EMU_PER_PX) * cs;
  const padRight = (txt.rIns / EMU_PER_PX) * cs;
  const padTop = (txt.tIns / EMU_PER_PX) * cs;
  const padBottom = (txt.bIns / EMU_PER_PX) * cs;
  const innerW = Math.max(0, sw - padLeft - padRight);
  const innerH = Math.max(0, sh - padTop - padBottom);
  if (innerW <= 0 || innerH <= 0) return;

  // A laid-out segment: measured text or a rasterized equation. `w` is the
  // advance width (px); math also carries baseline-relative ascent/descent.
  type Seg =
    | { kind: 'text'; text: string; font: string; color: string; w: number }
    | { kind: 'math'; render: MathRender; color: string; w: number; ascent: number; descent: number };
  // `leftInset` = px from padLeft to this line's left edge (paragraph left
  // margin, plus the first-line indent on a paragraph's first line). `availW` = the
  // width of the alignment region for this line (paraW, minus the first-line
  // indent on the first line). ECMA-376 §21.1.2.2.7 (marL/marR/indent).
  type Line = { segs: Seg[]; align: string; height: number; ascent: number; hasMath: boolean; leftInset: number; availW: number };

  // Font string + px size for a text run (math runs have no run-level font).
  const textFont = (run: Extract<import('./types.js').ShapeTextRun, { type: 'text' }>): { font: string; px: number } => {
    const size = run.size > 0 ? run.size : DEFAULT_FONT_SIZE;
    const px = size * PT_TO_PX * cs;
    const family = fontStackFor(run.fontFace);
    return { font: `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${px}px ${family}`, px };
  };

  // Real font ascent for a line's alphabetic baseline (used on lines that mix
  // text with an inline/display equation, where text and math share a baseline
  // = lineTop + ascent). Mirrors the pptx renderer: measure the rendered glyphs'
  // `actualBoundingBoxAscent` rather than the old flat 0.85×em heuristic, which
  // over-/under-estimated for CJK and tall fonts. `font` must already be a valid
  // CSS font string (measureText keys off `ctx.font`). Falls back to 0.85×px
  // when the metric is unavailable (0), matching the prior constant.
  const measuredAscent = (font: string, px: number): number => {
    const prev = ctx.font;
    ctx.font = font;
    const a = ctx.measureText('M').actualBoundingBoxAscent;
    ctx.font = prev;
    return a > 0 ? a : px * 0.85;
  };

  // Wrap each paragraph into lines (segments preserve run order).
  const wrap = txt.wrap !== 'none';
  const lines: Line[] = [];
  for (const p of txt.paragraphs) {
    const align = p.align || 'l';
    // ECMA-376 §21.1.2.2.7 direct paragraph indent (EMU → px, scaled by cs).
    // Mirrors the pptx renderer (marLPx/marRPx/indentPx + firstLineIndent).
    // Direct-attribute-only: xlsx text boxes have no lstStyle/level cascade, so
    // the spec's literal implied defaults (marL=347663, indent=−342900) are
    // deliberately NOT applied — there is no list-style tier to feed them, and
    // pptx's resolver leaves a plain bulletless paragraph at 0 too. Absent ⇒ 0.
    const marLpx = ((p.marL ?? 0) / EMU_PER_PX) * cs;
    const marRpx = ((p.marR ?? 0) / EMU_PER_PX) * cs;
    const indentPx = ((p.indent ?? 0) / EMU_PER_PX) * cs;
    // First-line indent eats into available width only when positive; a hanging
    // (negative) indent has no bullet gutter in xlsx, so it is clamped to 0.
    const firstLineIndent = Math.max(0, indentPx);
    const paraW = Math.max(0, innerW - marLpx - marRpx);
    // First line of the paragraph carries marLpx + firstLineIndent; continuation
    // lines carry only marLpx. Flipped to true on the first flush within the
    // paragraph (and on a display-math line, which is its own line).
    let firstLineDone = false;
    const lineLeftInset = () => (firstLineDone ? marLpx : marLpx + firstLineIndent);
    const lineAvailW = () => (firstLineDone ? paraW : paraW - firstLineIndent);
    let segs: Seg[] = [];
    let lineW = 0;
    let lineHeight = 0;
    let lineAscent = 0;
    let hasMath = false;
    // ECMA-376 §21.1.2.2.5 <a:lnSpc> + §21.1.2.1.3 normAutofit lnSpcReduction,
    // applied to a natural (design-floored) single-line height. Mirrors the pptx
    // renderer. `h` is the natural single line, so it is the correct pct base;
    // pts is an absolute per-line height (cs-scaled, like the cell/run px sizes).
    const applyLineSpacing = (h: number): number => {
      let out = h;
      if (p.spaceLine) {
        if (p.spaceLine.type === 'pct') out = out * (p.spaceLine.val / 100000);
        else out = p.spaceLine.val * PT_TO_PX * cs;
      }
      // normAutofit lnSpcReduction (§21.1.2.1.3): apply the STORED reduction
      // only, and ONLY to paragraphs with PERCENTAGE line spacing — the spec's
      // normative note reads "This attribute applies only to paragraphs with
      // percentage line spacing." So pct and the implicit single (= 100 %
      // percentage) get it; an absolute spcPts height does NOT. fontScale
      // font-shrink and spAutoFit shape-grow are runtime layout behaviors
      // intentionally out of scope (modeled but not applied) — the repo requires
      // explicit user approval for reverse-engineered autofit.
      if (txt.autoFit === 'norm' && txt.lnSpcReduction != null && p.spaceLine?.type !== 'pts') {
        out *= 1 - txt.lnSpcReduction;
      }
      return out;
    };
    const flushLine = () => {
      // An empty paragraph (no runs) or a blank line produced by a standalone /
      // trailing <a:br> contributes no text or math segment, so lineHeight is
      // still 0. ECMA-376 §21.1.2.1 / §21.1.2.2: such a line still reserves ONE
      // single-line height — as tall as a one-character line of the paragraph's
      // effective font. Without this the empty line reserved zero height, so the
      // block under-measured and vertical anchoring ('ctr'/'b') drifted. Mirror
      // the text-line formula (pxSize * 1.2, floored by the font's design line —
      // see the run sites) using the nearest preceding text size AND face in
      // this paragraph, falling back to the body default.
      if (lineHeight === 0) {
        const fallbackPx = (lastTextPt || DEFAULT_FONT_SIZE) * PT_TO_PX * cs;
        const designFloor = Math.max(
          intendedSingleLinePx(lastTextFace, fallbackPx),
          intendedSingleLinePx(lastTextFaceEa, fallbackPx),
        );
        lineHeight = Math.max(fallbackPx * 1.2, designFloor);
        // An empty line carries no segment, so this ascent is never consumed by
        // the draw pass (no text/math to place on the baseline); it is set only
        // to keep the Line shape consistent. Measure it the same way as text
        // runs — the nearest preceding face at the fallback size — rather than
        // the old 0.85×em constant.
        lineAscent = measuredAscent(`${fallbackPx}px ${fontStackFor(lastTextFace)}`, fallbackPx);
      }
      lineHeight = applyLineSpacing(lineHeight);
      lines.push({ segs, align, height: lineHeight, ascent: lineAscent, hasMath, leftInset: lineLeftInset(), availW: lineAvailW() });
      firstLineDone = true;
      segs = []; lineW = 0; lineHeight = 0; lineAscent = 0; hasMath = false;
    };
    // Nearest preceding text size (pt) in this paragraph — inline math with no
    // explicit rPr@sz inherits it (then falls back to the default).
    let lastTextPt = 0;
    // Nearest preceding text AUTHORED faces — used to floor an empty/blank
    // line's reserved single-line height by the tallest design line among the
    // declared latin / ea faces (intendedSingleLinePx), matching the text-run
    // floor below (cs is excluded from the line-box floor — see there).
    let lastTextFace: string | undefined;
    let lastTextFaceEa: string | undefined;

    for (const run of p.runs) {
      if (run.type === 'break') { flushLine(); continue; }

      if (run.type === 'math') {
        const render = mathRenders.get(run.nodes);
        if (!render) continue; // engine not supplied / conversion failed → skip
        const px = (run.fontSize ?? (lastTextPt || DEFAULT_FONT_SIZE)) * PT_TO_PX * cs;
        const w = render.widthEm * px;
        const ascent = render.ascentEm * px;
        const descent = render.descentEm * px;
        const color = run.color ?? '#000000';
        if (run.display) {
          // Block equation occupies its own line (centered per paragraph align).
          // It takes the paragraph left margin (marLpx) but NOT the first-line
          // indent — `indent` is a run-in indent for the first line of TEXT, not
          // for a block equation — so use marLpx/paraW regardless of line position.
          flushLine();
          // Apply the paragraph's line spacing to a display equation's own line
          // too (a block equation in a pct-spaced paragraph). ascent is kept
          // unchanged; the alphabetic-baseline draw distributes the extra leading.
          lines.push({ segs: [{ kind: 'math', render, color, w, ascent, descent }], align, height: applyLineSpacing(ascent + descent), ascent, hasMath: true, leftInset: marLpx, availW: paraW });
          firstLineDone = true;
          continue;
        }
        // Inline equation: treat as an atomic, non-breaking "word". Budget is
        // this line's available width (paraW, minus first-line indent on the
        // first line) rather than the full innerW.
        if (wrap && lineW + w > lineAvailW() && segs.length > 0) flushLine();
        segs.push({ kind: 'math', render, color, w, ascent, descent });
        lineW += w;
        lineHeight = Math.max(lineHeight, ascent + descent);
        lineAscent = Math.max(lineAscent, ascent);
        hasMath = true;
        continue;
      }

      // Text run.
      lastTextPt = run.size > 0 ? run.size : DEFAULT_FONT_SIZE;
      lastTextFace = run.fontFace;
      lastTextFaceEa = run.fontFaceEa;
      const { font, px: pxSize } = textFont(run);
      const color = run.color ?? '#000000';
      // Floor the natural single line (Excel's flat 1.2×em) by the AUTHORED
      // font's design line box (OS/2 win metrics, ECMA-376 §21.1.2.1.1) via
      // core's intendedSingleLinePx — same floor docx/pptx apply. It returns 0
      // for every untabled face (Calibri etc. stay on 1.2×em); a substituted
      // Meiryo (1.596×em) / Sakkal Majalla must measure to its taller design
      // line. Floor by the tallest of the LATIN and EAST-ASIAN faces (the common
      // Japanese encoding sets Meiryo only on `<a:ea>` while leaving `<a:latin>`
      // default, §21.1.2.3.1). `<a:cs>` is parsed (see fontFaceCs) but
      // deliberately NOT in this line-box floor: per the font-slot rules
      // (§21.1.2.3.1 / §17.3.2.26) the complex-script face renders ONLY
      // complex-script glyphs (Arabic/Hebrew/Thai), so an unconditional
      // line-box floor by cs would over-grow a run whose glyphs are Latin/CJK
      // (e.g. a Japanese run that merely also declares a tabled cs face). Getting
      // cs right needs per-glyph/per-script handling (deferred, like pptx's
      // per-glyph floor). Pass the authored names (the metric table keys on
      // them), NOT the fallback stack. FLOOR, not a replace; matches the docx
      // shape-text floor's max(latin, ea).
      const designFloor = Math.max(
        intendedSingleLinePx(run.fontFace, pxSize),
        intendedSingleLinePx(run.fontFaceEa, pxSize),
      );
      const singleLinePx = Math.max(pxSize * 1.2, designFloor);
      lineHeight = Math.max(lineHeight, singleLinePx);
      lineAscent = Math.max(lineAscent, measuredAscent(font, pxSize));
      ctx.font = font;
      // Defensive: a run's text may still contain a literal "\n".
      const pieces = run.text.split('\n');
      for (let s = 0; s < pieces.length; s++) {
        if (s > 0) flushLine();
        const piece = pieces[s];
        if (!piece) continue;
        if (!wrap) {
          // `measureAdvanceWithTrailingSpace` (not raw measureText): a piece may
          // end in whitespace whose advance Firefox/WebKit trim from measureText,
          // which would collapse an inter-word space at the piece edge. The core
          // helper restores it (pass-through on Chrome/skia); the draw pen
          // advances by this stored `w`, so measure==draw on every engine.
          const w = measureAdvanceWithTrailingSpace(ctx, piece);
          segs.push({ kind: 'text', text: piece, font, color, w });
          lineW += w;
          continue;
        }
        // Greedy character-level wrap (adequate for Latin + CJK). The wrap
        // budget is the CURRENT line's available width — paraW on continuation
        // lines, paraW − firstLineIndent on the paragraph's first line — not the
        // full innerW (ECMA-376 §21.1.2.2.7 marL/marR/indent).
        let buf = '';
        for (const ch of piece) {
          const candidate = buf + ch;
          const cw = ctx.measureText(candidate).width;
          if (lineW + cw > lineAvailW() && (buf.length > 0 || segs.length > 0)) {
            if (buf) {
              // A wrap buffer flushed here can end in a space; measure it
              // trailing-space-robustly so the space keeps its advance on
              // trimming engines (measure==draw with the draw pen below).
              const w = measureAdvanceWithTrailingSpace(ctx, buf);
              segs.push({ kind: 'text', text: buf, font, color, w });
              lineW += w;
            }
            flushLine();
            buf = ch;
            ctx.font = font;
            // Re-seed this continuation line with the same design-line-floored
            // single-line height as the run's first line (see singleLinePx above).
            lineHeight = Math.max(lineHeight, singleLinePx);
            lineAscent = Math.max(lineAscent, measuredAscent(font, pxSize));
          } else {
            buf = candidate;
          }
        }
        if (buf) {
          // Final buffer of the piece — same trailing-space-robust measure so a
          // word ending in a space keeps its advance (see above).
          const w = measureAdvanceWithTrailingSpace(ctx, buf);
          segs.push({ kind: 'text', text: buf, font, color, w });
          lineW += w;
        }
      }
    }
    flushLine();
  }

  // Total text block height
  const blockH = lines.reduce((s, l) => s + l.height, 0);

  // Vertical anchor — ECMA-376 §20.1.7.2 <a:bodyPr anchor>.
  // For 'ctr' we intentionally skip Math.max(0,...) so the block stays
  // visually centered even when blockH exceeds innerH (text clips at edge).
  let y0 = padTop;
  if (txt.anchor === 'ctr') y0 = padTop + (innerH - blockH) / 2;
  else if (txt.anchor === 'b') y0 = padTop + Math.max(0, innerH - blockH);

  let lineTop = y0;
  for (const line of lines) {
    const totalW = line.segs.reduce((s, seg) => s + seg.w, 0);
    // Per-line region: the left edge is padLeft + the paragraph's left inset
    // (marL, plus first-line indent on the first line), and alignment happens
    // within the line's available width (paraW). ECMA-376 §21.1.2.2.7.
    const base = padLeft + line.leftInset;
    let x = base;
    if (line.align === 'ctr') x = base + Math.max(0, (line.availW - totalW) / 2);
    else if (line.align === 'r') x = base + Math.max(0, line.availW - totalW);

    if (line.hasMath) {
      // A line containing an equation aligns text AND the math raster to a
      // shared alphabetic baseline (= lineTop + ascent), so the equation sits
      // on the same baseline as adjacent text. (Pure-text lines keep the
      // simpler 'middle' baseline below.)
      ctx.textBaseline = 'alphabetic';
      const baseline = lineTop + line.ascent;
      for (const seg of line.segs) {
        if (seg.kind === 'text') {
          ctx.font = seg.font;
          ctx.fillStyle = seg.color;
          ctx.fillText(seg.text, x, baseline);
        } else {
          const img = tintedMathImage(seg.render, seg.color);
          ctx.drawImage(img, x, baseline - seg.ascent, seg.w, seg.ascent + seg.descent);
        }
        x += seg.w;
      }
    } else {
      ctx.textBaseline = 'middle';
      const drawY = lineTop + line.height / 2;
      for (const seg of line.segs) {
        if (seg.kind === 'text') {
          ctx.font = seg.font;
          ctx.fillStyle = seg.color;
          ctx.fillText(seg.text, x, drawY);
        }
        x += seg.w;
      }
    }
    lineTop += line.height;
  }
}

function fillAndStroke(ctx: CanvasRenderingContext2D, shape: ShapeInfo): void {
  if (shape.fillColor) {
    ctx.fillStyle = shape.fillColor;
    ctx.fill();
  }
  if (shape.strokeColor && shape.strokeWidth > 0) {
    ctx.strokeStyle = shape.strokeColor;
    ctx.lineWidth = Math.max(0.5, shape.strokeWidth / EMU_PER_PX);
    ctx.stroke();
  }
}

// ────────────────────────────────────────────────────────────────
// Border drawing
// ────────────────────────────────────────────────────────────────
/**
 * Overlay any CF-rule border edges on top of the cell's base border. CF
 * borders win per-edge where they set a style (e.g. a red left+right for a
 * "today" column marker replaces the underlying edge only, leaving top/bottom
 * from the base style intact).
 */
/**
 * Resolve the outer border of a merged range. Excel keeps each constituent
 * cell's own style; the merged rectangle's outer edges come from the cells
 * on those edges (e.g. the right edge from the rightmost column's `right`),
 * not from the anchor cell alone. Without this the right border of an
 * `E2:F2` merge — stored on F2 — goes missing.
 */
function resolveMergeBorder(
  anchorBorder: Border,
  anchorRow: number,
  anchorCol: number,
  rightCol: number,
  bottomRow: number,
  cellMap: Map<string, Cell>,
  styles: Styles,
): Border {
  if (rightCol === anchorCol && bottomRow === anchorRow) return anchorBorder;
  const edgeBorder = (r: number, c: number): Border | null => {
    if (r === anchorRow && c === anchorCol) return null;
    const cell = cellMap.get(`${r}:${c}`);
    if (!cell) return null;
    return resolveXf(styles, cell.styleIndex ?? 0).border;
  };
  const rightB  = edgeBorder(anchorRow, rightCol);
  const bottomB = edgeBorder(bottomRow, anchorCol);
  const cornerB = edgeBorder(bottomRow, rightCol);
  const pick = (primary: BorderEdge | null | undefined, ...rest: Array<BorderEdge | null | undefined>): BorderEdge | null => {
    if (primary?.style) return primary;
    for (const r of rest) if (r?.style) return r;
    return primary ?? null;
  };
  return {
    left:         anchorBorder.left,
    top:          anchorBorder.top,
    right:        pick(rightB?.right,   cornerB?.right,   anchorBorder.right),
    bottom:       pick(bottomB?.bottom, cornerB?.bottom,  anchorBorder.bottom),
    diagonalUp:   anchorBorder.diagonalUp ?? null,
    diagonalDown: anchorBorder.diagonalDown ?? null,
  };
}

function mergeBorders(base: Border, overlay: Border | undefined): Border {
  if (!overlay) return base;
  const pick = (a: BorderEdge | null | undefined, b: BorderEdge | null | undefined): BorderEdge | null =>
    (b && b.style) ? b : (a ?? null);
  return {
    left:         pick(base.left,         overlay.left),
    right:        pick(base.right,        overlay.right),
    top:          pick(base.top,          overlay.top),
    bottom:       pick(base.bottom,       overlay.bottom),
    diagonalUp:   pick(base.diagonalUp,   overlay.diagonalUp),
    diagonalDown: pick(base.diagonalDown, overlay.diagonalDown),
  };
}

function renderBorder(
  ctx: CanvasRenderingContext2D,
  border: Border,
  x: number, y: number, w: number, h: number,
  dpr = 1,
): void {
  type EdgeRef = {
    edge: BorderEdge | null | undefined;
    x1: number; y1: number; x2: number; y2: number;
    /** 'h' = horizontal (top/bottom), 'v' = vertical (left/right),
     *  'd' = diagonal (no doubling support). */
    kind: 'h' | 'v' | 'd';
  };
  const edges: EdgeRef[] = [
    { edge: border.top,         x1: x,     y1: y,     x2: x + w, y2: y,     kind: 'h' },
    { edge: border.bottom,      x1: x,     y1: y + h, x2: x + w, y2: y + h, kind: 'h' },
    { edge: border.left,        x1: x,     y1: y,     x2: x,     y2: y + h, kind: 'v' },
    { edge: border.right,       x1: x + w, y1: y,     x2: x + w, y2: y + h, kind: 'v' },
    { edge: border.diagonalUp,  x1: x,     y1: y + h, x2: x + w, y2: y,     kind: 'd' },
    { edge: border.diagonalDown,x1: x,     y1: y,     x2: x + w, y2: y + h, kind: 'd' },
  ];
  for (const { edge, x1, y1, x2, y2, kind } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    const color = edge.color ? hexToRgba(edge.color) : '#000000';
    // ECMA-376 §18.18.3 ST_BorderStyle "double": two parallel thin lines
    // with a small gap. Drawn as a 1-px line on either side of the cell
    // edge so the pair reads as a ~3-px-thick double rule, matching Excel.
    // The outer line extends past the cell corners by `off` so adjacent
    // doubled edges close cleanly; the inner line is *shortened* by `off`
    // on each end so the two perpendicular inner lines meet exactly at the
    // inner corner (without the trim they cross past the corner forming a
    // small "+" overhang). Diagonals are handled separately further down.
    if (edge.style === 'double' && kind === 'd') {
      // Two parallel diagonal lines offset perpendicular to the diagonal
      // direction. Excel renders <diagonal style="double"/> the same way it
      // does horizontal/vertical doubles — two thin lines with a small gap.
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const dx = x2 - x1;
      const dy = y2 - y1;
      const len = Math.hypot(dx, dy);
      // Perpendicular unit vector. ±off shifts the line ~1 logical px to either side.
      const off = 1;
      const px = (-dy / len) * off;
      const py = (dx / len) * off;
      ctx.beginPath();
      ctx.moveTo(x1 + px, y1 + py); ctx.lineTo(x2 + px, y2 + py);
      ctx.moveTo(x1 - px, y1 - py); ctx.lineTo(x2 - px, y2 - py);
      ctx.stroke();
      continue;
    }
    if (edge.style === 'double' && kind !== 'd') {
      // NB: core's `fillDoubleBorder` (used by docx) is intentionally NOT used
      // here. Excel's double is two symmetric 1px strokes at ±1px with corner
      // extension/trim so a double BOX closes cleanly; the core fill helper's
      // floored-thirds band has a +0.5-device centring bias and no corner
      // handling, which measured ~0.3% further from the Excel references
      // (sample-11) when trialled. These are different conventions, not
      // duplicated math — keep the Excel-faithful stroke model.
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.setLineDash([]);
      const off = 1;
      ctx.beginPath();
      if (kind === 'h') {
        const isTop = y1 === y;
        // Outer line extends past the corners by `off` so adjacent doubled
        // edges close; the inner line is shortened by `off` on each end so
        // perpendicular inner lines meet exactly at the inner corner. In the
        // two-pass order each side draws its own double fully (no fill erases
        // it), so no inherited-redraw "swap" is needed.
        const outerY = isTop ? y - off : y + h + off;
        const innerY = isTop ? y + off : y + h - off;
        ctx.moveTo(x - off, outerY);   ctx.lineTo(x + w + off, outerY);
        ctx.moveTo(x + off, innerY);   ctx.lineTo(x + w - off, innerY);
      } else {
        const isLeft = x1 === x;
        const outerX = isLeft ? x - off : x + w + off;
        const innerX = isLeft ? x + off : x + w - off;
        ctx.moveTo(outerX, y - off);   ctx.lineTo(outerX, y + h + off);
        ctx.moveTo(innerX, y + off);   ctx.lineTo(innerX, y + h - off);
      }
      ctx.stroke();
      continue;
    }
    ctx.beginPath();
    ctx.strokeStyle = color;
    // Logical-px width (thin=1, medium=2, thick=3). ctx.scale(dpr,dpr) scales it
    // to device px, so a thin border is 2 device px at dpr=2 — matching Excel's
    // measured on-screen weight. Do NOT divide by dpr (that halved it to 1 px).
    const lw = borderStyleWidth(edge.style);
    ctx.lineWidth = lw;
    const dash = borderStyleDash(edge.style);
    ctx.setLineDash(dash);
    // Crispness snap (see crispOffset): snap the stroke to the nearest crisp
    // device position derived from its own coordinate — the x of a vertical
    // edge, the y of a horizontal edge. An odd device-width stroke lands on a
    // pixel midpoint; an even device width snaps to an integer boundary (0 when
    // already aligned). Diagonals can't be pixel-aligned, so no offset.
    const dpx = kind === 'v' ? crispOffset(x1, lw, dpr) : 0;
    const dpy = kind === 'h' ? crispOffset(y1, lw, dpr) : 0;
    ctx.moveTo(x1 + dpx, y1 + dpy);
    ctx.lineTo(x2 + dpx, y2 + dpy);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function borderStyleWidth(style: string): number {
  // Widths follow Excel's pt convention so medium and thick read as visibly
  // distinct from thin (and from each other) — thin=1pt, medium=2pt, thick=3pt.
  // Earlier values (1.5 / 2) compressed medium and thick to roughly the same
  // 2-row antialiased band, which made e.g. sample-27 row 9 appear to have
  // top (medium) and bottom (thick) of equal weight.
  switch (style) {
    case 'thick': return 3;
    case 'medium': case 'mediumDashed': case 'mediumDashDot': case 'mediumDashDotDot': case 'slantDashDot': return 2;
    case 'hair': return 0.5;
    default: return 1;
  }
}

/**
 * ECMA-376 §18.18.3 ST_BorderStyle dash families → a static-pixel `setLineDash`
 * pattern. Thin wrapper over core's shared `xlsxBorderDashArray` (which owns the
 * §18.18.3 relative table); Excel cell borders use a constant pixel cadence
 * regardless of the (sub-pixel) hairline width, so unlike docx/pptx this does
 * NOT scale with the stroked width. Returns `[]` for solid styles.
 */
function borderStyleDash(style: string): number[] {
  return xlsxBorderDashArray(style);
}

/**
 * Visual precedence per ECMA-376 ST_BorderStyle. Higher = stronger / more
 * visually prominent. Excel's behaviour at a shared edge between two adjacent
 * cells that both define a border is to render the stronger one — this lets
 * the renderer pick a single edge to draw instead of stacking two strokes
 * (which compounds antialiasing and visually thickens the line).
 */
function borderPrecedence(style: string | null | undefined): number {
  switch (style) {
    case 'double': return 13;
    case 'thick': return 12;
    case 'medium': return 11;
    case 'mediumDashed': return 10;
    case 'mediumDashDot': return 9;
    case 'slantDashDot': return 8;
    case 'mediumDashDotDot': return 7;
    case 'thin': return 6;
    case 'dashed': return 5;
    case 'dashDot': return 4;
    case 'dashDotDot': return 3;
    case 'dotted': return 2;
    case 'hair': return 1;
    default: return 0;
  }
}

function pickStrongerEdge(
  a: BorderEdge | null | undefined,
  b: BorderEdge | null | undefined,
): BorderEdge | null {
  const aP = borderPrecedence(a?.style);
  const bP = borderPrecedence(b?.style);
  if (aP === 0 && bP === 0) return null;
  if (aP >= bP) return a ?? null;
  return b ?? null;
}

// ═══════════════════════════════════════════════════════════════════════════
// Chart rendering — delegated to @silurus/ooxml-core's unified renderer.
//
// The parser now emits the canonical `ChartModel` directly (see the Rust
// `From<ChartData> for ChartModel`), so there is no TS adapter here anymore —
// `anchor.chart` is passed straight to `renderChart`.
// ═══════════════════════════════════════════════════════════════════════════

// ── renderCharts ────────────────────────────────────────────────────────────

function renderCharts(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of ws.charts) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const cw = shX2 - shX1;
    const ch = shY2 - shY1;
    if (cw <= 0 || ch <= 0) continue;

    const cx = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const cy = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (cx + cw < scrollAreaX || cx > scrollAreaX + scrollAreaW) continue;
    if (cy + ch < scrollAreaY || cy > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    // XLSX natural rendering is device-px at 96 DPI where 1pt = 4/3 px. Scale
    // that by `cs` so OOXML-specified font sizes (title/axes) scale with zoom.
    const ptToPx = PT_TO_PX * cs;
    // `anchor.chart` is already the canonical ChartModel emitted by the Rust
    // parser (`ooxml_common::chart::ChartModel`) — the former `adaptChartData`
    // default/mapping logic now lives in the parser's `From<ChartData>`.
    renderChart(ctx, anchor.chart, { x: cx, y: cy, w: cw, h: ch }, ptToPx);
    ctx.restore();
  }
}

// ── renderSlicers ───────────────────────────────────────────────────────────
//
// Office 2010+ pivot / table slicer. We don't own a slicer engine, so this
// renders a static button bank: header with the slicer caption, then one
// button per saved item using the selection flags from the slicerCache. The
// visual language (pale blue outline, white "selected" buttons on a darker
// background, gray "deselected" buttons) intentionally mirrors Excel's
// default slicer style — the workbook may ship a custom `slicerStyle` but
// rendering that is deferred (the built-in look is already recognisable).

const SLICER_HEADER_FONT = '600 12px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_ITEM_FONT   = '11px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_BG           = '#FFFFFF';
const SLICER_BORDER       = '#BFBFBF';
const SLICER_HEADER_BG    = '#F2F2F2';
const SLICER_HEADER_FG    = '#404040';
const SLICER_ITEM_SEL_BG  = '#FFFFFF';
const SLICER_ITEM_SEL_FG  = '#000000';
const SLICER_ITEM_SEL_BD  = '#A5A5A5';
const SLICER_ITEM_OFF_BG  = '#E7E6E6';
const SLICER_ITEM_OFF_FG  = '#A6A6A6';
const SLICER_ITEM_OFF_BD  = '#C6C6C6';

function renderSlicers(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const slicers = ws.slicers;
  if (!slicers) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of slicers) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const w = shX2 - shX1;
    const h = shY2 - shY1;
    if (w <= 0 || h <= 0) continue;

    const x = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const y = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (x + w < scrollAreaX || x > scrollAreaX + scrollAreaW) continue;
    if (y + h < scrollAreaY || y > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    drawSlicerFrame(ctx, anchor.caption, anchor.items, x, y, w, h, cs);

    ctx.restore();
  }
}

function drawSlicerFrame(
  ctx: CanvasRenderingContext2D,
  caption: string,
  items: SlicerItem[],
  x: number,
  y: number,
  w: number,
  h: number,
  cs: number,
): void {
  // Outer frame (white with a soft gray hairline).
  ctx.fillStyle = SLICER_BG;
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = SLICER_BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);

  // Header band with caption.
  const headerH = Math.max(20 * cs, 14);
  ctx.fillStyle = SLICER_HEADER_BG;
  ctx.fillRect(x + 1, y + 1, w - 2, headerH);
  ctx.fillStyle = SLICER_HEADER_FG;
  ctx.font = scaleFont(SLICER_HEADER_FONT, cs);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'left';
  const headerPad = 6 * cs;
  drawClippedText(ctx, caption, x + headerPad, y + headerH / 2 + 1, w - 2 * headerPad);

  // Item buttons. Items expand to fill the available height (up to a
  // minimum button size) and are clipped by the slicer rect. We don't
  // implement scroll arrows because this renderer is non-interactive.
  if (items.length === 0) return;
  const gap = Math.max(1, Math.round(2 * cs));
  const innerPad = 4 * cs;
  const listX = x + innerPad;
  const listY = y + headerH + innerPad;
  const listW = w - 2 * innerPad;
  const listH = h - headerH - 2 * innerPad;
  if (listW <= 0 || listH <= 0) return;

  // Prefer Excel's rough row height (~20 sheet-px) but compress if the
  // slicer is shallow so at least the first items fit.
  const preferredItemH = Math.max(18 * cs, 16);
  const maxVisibleByH = Math.max(1, Math.floor((listH + gap) / (preferredItemH + gap)));
  const visible = Math.min(items.length, maxVisibleByH);
  const itemH = Math.min(preferredItemH, (listH - gap * (visible - 1)) / visible);
  if (itemH <= 0) return;

  ctx.font = scaleFont(SLICER_ITEM_FONT, cs);
  const itemPad = 8 * cs;
  for (let i = 0; i < visible; i++) {
    const item = items[i];
    const iy = listY + i * (itemH + gap);
    const selected = item.selected;
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_BG : SLICER_ITEM_OFF_BG;
    ctx.fillRect(listX, iy, listW, itemH);
    ctx.strokeStyle = selected ? SLICER_ITEM_SEL_BD : SLICER_ITEM_OFF_BD;
    ctx.lineWidth = 1;
    ctx.strokeRect(listX + 0.5, iy + 0.5, listW - 1, itemH - 1);
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_FG : SLICER_ITEM_OFF_FG;
    drawClippedText(ctx, item.name, listX + itemPad, iy + itemH / 2 + 1, listW - 2 * itemPad);
  }
}

function scaleFont(css: string, cs: number): string {
  // Re-scale the leading `<size>px` token by `cs`. Safe fallback: leave the
  // string as-is so the slicer remains readable when parsing fails.
  return css.replace(/(\d+(?:\.\d+)?)px/, (_, n) => `${Math.round(Number(n) * cs)}px`);
}

function drawClippedText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  maxWidth: number,
): void {
  if (maxWidth <= 0) return;
  let s = text;
  if (ctx.measureText(s).width > maxWidth) {
    const ellipsis = '…';
    while (s.length > 0 && ctx.measureText(s + ellipsis).width > maxWidth) {
      s = s.slice(0, -1);
    }
    s = s.length > 0 ? s + ellipsis : '';
  }
  ctx.fillText(s, x, y);
}
