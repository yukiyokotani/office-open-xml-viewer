import { canvasFontString } from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import type { GlyphOwnershipRequest, GlyphRasterComparisonRequest } from './layout/text.js';

type RasterContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type RasterCanvas = HTMLCanvasElement | OffscreenCanvas;

/** Probe em in CSS pixels. Large enough that synthetic emboldening and
 * outline differences between faces reach whole device pixels; small enough
 * to keep each proof to a few hundred kilobytes of pixel reads. */
const PROBE_EM_PX = 64;
/** Upper bound on scalars rasterized per proof (callers pass a short probe). */
const PROBE_SCALARS = 8;

function createRasterCanvas(width: number, height: number): RasterCanvas | null {
  if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(width, height);
  if (typeof document !== 'undefined' && document?.createElement) {
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    return canvas;
  }
  return null;
}

// One realm-local surface, resized per proof. Proofs are synchronous and
// never re-entered, so sharing it is safe and avoids per-proof allocation.
let shared: { canvas: RasterCanvas; context: RasterContext } | null | undefined;

function sharedContext(width: number, height: number): RasterContext | null {
  if (shared === undefined) {
    const canvas = createRasterCanvas(1, 1);
    const context = canvas?.getContext('2d', { willReadFrequently: true }) as RasterContext | null | undefined;
    shared = canvas && context ? { canvas, context } : null;
  }
  if (!shared) return null;
  if (shared.canvas.width < width || shared.canvas.height !== height) {
    shared.canvas.width = Math.max(width, shared.canvas.width);
    shared.canvas.height = height;
  }
  return shared.context;
}

/** Owned scratch surfaces may be resized between calls (which resets Canvas
 * state). Set/read/restore the requested base for each actual native operation. */
function withDirection<T>(context: RasterContext, base: 'ltr' | 'rtl' | undefined, action: () => T): T {
  if (base === undefined) return action();
  const previous = context.direction;
  const previousAlign = context.textAlign;
  try {
    context.direction = base;
    context.textAlign = 'left';
    if (context.direction !== base || context.textAlign !== 'left') {
      throw new TypeError('Canvas did not honor directed glyph proof and left anchor');
    }
    return action();
  } finally { context.direction = previous; context.textAlign = previousAlign; }
}

function rasterize(context: RasterContext, font: string, text: string, width: number, height: number, base?: 'ltr' | 'rtl'): Uint8ClampedArray {
  context.clearRect(0, 0, width, height);
  context.font = font;
  context.textBaseline = 'alphabetic';
  context.fillStyle = '#000';
  withDirection(context, base, () => context.fillText(text, PROBE_EM_PX / 2, Math.round(height * 0.7)));
  return context.getImageData(0, 0, width, height).data;
}

/**
 * Proof that a native CSS route paints the same face as one loaded exact
 * local() face (issue #1848). The comparison family is a private alias that
 * holds exactly one FontFace, so CSS Fonts Level 4 §5.2 font matching for any
 * weight/style tuple must select that face, and the browser synthesizes a
 * missing bold/oblique from it. When the native route rasterizes the probe
 * glyphs pixel-identically, it paints that same face with the same synthesis;
 * its vertical tables are therefore the loaded face's. Any difference (a
 * genuine styled face, a different file, glyphs from a fallback face) is no
 * proof. Both rasters come from this realm's Canvas and FontFaceSet, so main
 * thread and worker evaluate the same fonts.
 */
export function sameGlyphRaster(request: Readonly<GlyphRasterComparisonRequest>): boolean {
  const probe = [...request.text].filter((scalar) => !/\s/u.test(scalar)).slice(0, PROBE_SCALARS).join('');
  if (!probe) return false;
  const nativeFont = canvasFontString(request.route, PROBE_EM_PX, request.weight, request.style);
  const faceFont = `${request.style} ${request.weight} ${PROBE_EM_PX}px ${quoteCssFontFamily(request.faceFamily)}`;
  const measure = sharedContext(1, PROBE_EM_PX * 2);
  if (!measure) return false;
  measure.font = nativeFont;
  const nativeWidth = withDirection(measure, request.baseDirection, () => measure.measureText(probe).width);
  measure.font = faceFont;
  const faceWidth = withDirection(measure, request.baseDirection, () => measure.measureText(probe).width);
  if (!(Number.isFinite(nativeWidth) && nativeWidth > 0) || nativeWidth !== faceWidth) return false;
  const width = Math.ceil(nativeWidth + PROBE_EM_PX);
  const height = PROBE_EM_PX * 2;
  const context = sharedContext(width, height);
  if (!context) return false;
  try {
    const native = rasterize(context, nativeFont, probe, width, height, request.baseDirection);
    const face = rasterize(context, faceFont, probe, width, height, request.baseDirection);
    let ink = false;
    for (let index = 0; index < native.length; index++) {
      if (native[index] !== face[index]) return false;
      if (native[index] !== 0) ink = true;
    }
    return ink;
  } catch {
    // An unreadable raster is no proof; the caller keeps its Canvas box.
    return false;
  }
}

// Bounded per realm: a grapheme's ownership by one registration never changes.
const ownership = new Map<string, boolean>();
const OWNERSHIP_LIMIT = 16384;
/** Longer graphemes (long combining sequences) are proved but not cached, so
 * the realm-wide cache stays bounded in bytes as well as entries. */
const OWNERSHIP_KEY_UNITS = 16;
let graphemes: Intl.Segmenter | null | undefined;

/**
 * Glyph ownership through a size-adjust:0% registration of the exact source
 * (CSS Fonts 5 §4.4 scales outlines and advances, not cmap; the #1842 coverage
 * probe): a grapheme the source owns shapes with zero advance, while one it
 * lacks falls back to another face with a positive advance. Raster identity
 * proves which face a route selects; this proves the span's own glyphs come
 * from that face rather than a fallback (CSS Fonts 4 §5.2 permits fallback
 * from the selected face).
 */
export function ownsGlyphs(request: Readonly<GlyphOwnershipRequest>): boolean {
  const text = request.text;
  if (!text) return false;
  if (graphemes === undefined) {
    graphemes = typeof Intl !== 'undefined' && 'Segmenter' in Intl
      ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : null;
  }
  const parts = graphemes ? Array.from(graphemes.segment(text), (part) => part.segment) : [...text];
  const font = `${request.style} ${request.weight} 100px ${quoteCssFontFamily(request.coverageProbeFamily)}`;
  let context: RasterContext | null | undefined;
  for (const part of parts) {
    const key = request.baseDirection === undefined ? `${font}\u0000${part}`
      : `directed\u0000${JSON.stringify([font, part, request.baseDirection])}`;
    let owned = ownership.get(key);
    if (owned === undefined) {
      context ??= sharedContext(1, PROBE_EM_PX * 2);
      if (!context) return false;
      context.font = font;
      // Owned glyphs keep zero advance AND zero ink (the outlines scale too);
      // a fallback, even a zero-advance combining mark, keeps positive ink.
      const metrics = request.baseDirection === undefined
        ? context.measureText(part)
        : withDirection(context, request.baseDirection, () => context!.measureText(part));
      const inkWidth = (metrics.actualBoundingBoxLeft ?? 0) + (metrics.actualBoundingBoxRight ?? 0);
      const inkHeight = (metrics.actualBoundingBoxAscent ?? 0) + (metrics.actualBoundingBoxDescent ?? 0);
      owned = metrics.width === 0 && Number.isFinite(inkWidth) && Number.isFinite(inkHeight)
        && inkWidth <= 0 && inkHeight <= 0;
      if (part.length <= OWNERSHIP_KEY_UNITS) {
        if (ownership.size >= OWNERSHIP_LIMIT) ownership.clear();
        ownership.set(key, owned);
      }
    }
    if (!owned) return false;
  }
  return true;
}
