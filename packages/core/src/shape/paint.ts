import type { Fill, GradientFill, PatternFill, Stroke } from '../types/common';
import { buildPatternBitmap } from './pattern-bitmaps';
import { drawingmlLineDashArray, shapeStrokeDashArray } from '../draw/dash';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';

const MAX_GRADIENT_TILE_EDGE = 512;

function tiledGradient(
  fill: GradientFill,
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  shapeRotationDeg: number,
): CanvasPattern | null {
  const tile = fill.tileRect;
  if (!tile) return null;
  // An explicitly authored all-zero CT_RelativeRect is the complete fill
  // region. It has exactly one tile, so allocating two auxiliary canvases and
  // a repeating pattern would change no pixels and only add per-shape work.
  if ((tile.l ?? 0) === 0 && (tile.t ?? 0) === 0
      && (tile.r ?? 0) === 0 && (tile.b ?? 0) === 0) return null;
  const tileX = x + w * (tile.l ?? 0);
  const tileY = y + h * (tile.t ?? 0);
  const tileW = w * (1 - (tile.l ?? 0) - (tile.r ?? 0));
  const tileH = h * (1 - (tile.t ?? 0) - (tile.b ?? 0));
  if (!Number.isFinite(tileW) || !Number.isFinite(tileH)
      || Math.abs(tileW) < 1e-9 || Math.abs(tileH) < 1e-9) return null;

  const scale = Math.min(
    1,
    MAX_GRADIENT_TILE_EDGE / Math.abs(tileW),
    MAX_GRADIENT_TILE_EDGE / Math.abs(tileH),
  );
  const baseW = Math.max(1, Math.ceil(Math.abs(tileW) * scale));
  const baseH = Math.max(1, Math.ceil(Math.abs(tileH) * scale));
  const base = createAuxCanvasForContext(ctx, baseW, baseH);
  const baseCtx = base?.getContext('2d');
  if (!base || !baseCtx) return null;
  const basePaint = resolveFill(
    { ...fill, tileRect: undefined, flip: undefined },
    baseCtx as CanvasRenderingContext2D,
    0,
    0,
    baseW,
    baseH,
    shapeRotationDeg,
  );
  if (!basePaint) return null;
  baseCtx.fillStyle = basePaint;
  baseCtx.fillRect(0, 0, baseW, baseH);

  const flipX = fill.flip === 'x' || fill.flip === 'xy';
  const flipY = fill.flip === 'y' || fill.flip === 'xy';
  let patternSource = base;
  if (flipX || flipY) {
    const repeatW = baseW * (flipX ? 2 : 1);
    const repeatH = baseH * (flipY ? 2 : 1);
    const repeat = createAuxCanvasForContext(ctx, repeatW, repeatH);
    const repeatCtx = repeat?.getContext('2d');
    if (!repeat || !repeatCtx) return null;
    for (let row = 0; row < (flipY ? 2 : 1); row += 1) {
      for (let col = 0; col < (flipX ? 2 : 1); col += 1) {
        repeatCtx.save();
        repeatCtx.translate(col * baseW, row * baseH);
        repeatCtx.scale(col === 1 ? -1 : 1, row === 1 ? -1 : 1);
        repeatCtx.drawImage(base, col === 1 ? -baseW : 0, row === 1 ? -baseH : 0);
        repeatCtx.restore();
      }
    }
    patternSource = repeat;
  }
  const pattern = ctx.createPattern(patternSource, 'repeat');
  if (!pattern || typeof pattern.setTransform !== 'function') return null;
  pattern.setTransform({
    a: tileW / baseW,
    b: 0,
    c: 0,
    d: tileH / baseH,
    e: tileX,
    f: tileY,
  });
  return pattern;
}

/**
 * Convert a 6- or 8-char hex colour to a CSS `rgba()` string.
 * 8-char hex encodes alpha in the last two chars (RRGGBBAA).
 * `alpha` applies to 6-char hex; ignored for 8-char.
 * A leading `#` is tolerated (`#RRGGBB` and `RRGGBB` both work).
 */
export function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  const a = h.length >= 8 ? parseInt(h.slice(6, 8), 16) / 255 : alpha;
  return `rgba(${r},${g},${b},${a})`;
}

function colorCanProduceVisiblePixels(color: string): boolean {
  const normalized = color.startsWith('#') ? color.slice(1) : color;
  if (normalized.length < 8) return true;
  const alpha = Number.parseInt(normalized.slice(6, 8), 16);
  // Parsed OOXML colours are well formed. Treat malformed hand-authored
  // public-model input conservatively as visible rather than silently
  // suppressing paint or changing a chart-style role.
  return !Number.isFinite(alpha) || alpha !== 0;
}

/**
 * Static visibility bound for one DrawingML fill recipe. `true` means the
 * recipe may produce visible pixels; image contents themselves remain opaque
 * to this layer. A finite non-positive `a:alphaModFix` is nevertheless known
 * to make every image pixel transparent, so it must not trigger image decode,
 * resource charging, or callout-style selection.
 */
export function fillCanProduceVisiblePixels(fill: Fill | null | undefined): boolean {
  if (!fill || fill.fillType === 'none') return false;
  if (fill.fillType === 'solid') return colorCanProduceVisiblePixels(fill.color);
  if (fill.fillType === 'gradient') {
    return fill.stops.some(stop => colorCanProduceVisiblePixels(stop.color));
  }
  if (fill.fillType === 'pattern') {
    return colorCanProduceVisiblePixels(fill.fg) || colorCanProduceVisiblePixels(fill.bg);
  }
  return fill.alpha == null || !Number.isFinite(fill.alpha) || fill.alpha > 0;
}

/**
 * Rec.601 perceptual luma (`0.299·R + 0.587·G + 0.114·B`) of a colour, on the
 * 0–255 scale. Accepts a 6- or 8-char hex; a leading `#` is tolerated and the
 * alpha byte (if present) is ignored, matching {@link hexToRgba}'s hex
 * normalisation.
 */
export function relativeLuma(hex: string): number {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/**
 * Pick black or white text for legibility against a background colour. The
 * mid-gray threshold (128) on the Rec.601 luma splits light vs dark: a dark
 * background ⇒ white text, otherwise black. `bgHex=null` (no background ⇒ page
 * white) ⇒ black text. The black/white pick is implementation-defined — no
 * normative algorithm exists (ECMA-376 §17.3.2.6 `w:color="auto"` only says the
 * consumer chooses "an appropriate color based on the background").
 */
export function autoContrastColor(bgHex: string | null): '#000000' | '#FFFFFF' {
  if (!bgHex) return '#000000';
  return relativeLuma(bgHex) < 128 ? '#FFFFFF' : '#000000';
}

/**
 * Resolve a Fill to a CanvasRenderingContext2D-compatible paint.
 * Gradients require pixel bounds (x, y, w, h) to construct the CanvasGradient.
 * Returns null for noFill.
 */
export function resolveFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  shapeRotationDeg = 0,
): string | CanvasGradient | CanvasPattern | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'pattern') {
    return resolvePatternFill(fill, ctx);
  }
  if (fill.fillType === 'gradient') {
    const stops = fill.stops;
    if (stops.length === 0) return null;
    if (stops.length === 1) return hexToRgba(stops[0].color);

    const repeated = tiledGradient(fill, ctx, x, y, w, h, shapeRotationDeg);
    if (repeated) return repeated;

    let gradient: CanvasGradient;
    const tile = fill.tileRect;
    const tileX = x + w * (tile?.l ?? 0);
    const tileY = y + h * (tile?.t ?? 0);
    const tileW = w * (1 - (tile?.l ?? 0) - (tile?.r ?? 0));
    const tileH = h * (1 - (tile?.t ?? 0) - (tile?.b ?? 0));
    if (fill.gradType === 'radial') {
      // §20.1.8.31: fillToRect is the center-shade (focus) rectangle inside
      // the gradient tile. Canvas has a point focus rather than a rectangular
      // focus, so use its authored centre and retain the full rectangle on the
      // public model for richer hosts.
      const focus = fill.fillToRect;
      const focusX = tileX + tileW * (focus?.l ?? 0);
      const focusY = tileY + tileH * (focus?.t ?? 0);
      const focusW = tileW * (1 - (focus?.l ?? 0) - (focus?.r ?? 0));
      const focusH = tileH * (1 - (focus?.t ?? 0) - (focus?.b ?? 0));
      const cx = focusX + focusW / 2;
      const cy = focusY + focusH / 2;
      const rx = Math.max(Math.abs(cx - tileX), Math.abs(tileX + tileW - cx));
      const ry = Math.max(Math.abs(cy - tileY), Math.abs(tileY + tileH - cy));
      const r = fill.path === 'rect'
        ? Math.max(rx, ry)
        : Math.sqrt(rx * rx + ry * ry);
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(r, 1e-9));
    } else {
      const authoredAngle = fill.rotWithShape === false
        ? fill.angle - shapeRotationDeg
        : fill.angle;
      const rad = (authoredAngle * Math.PI) / 180;
      let dx = Math.cos(rad);
      let dy = Math.sin(rad);
      if (fill.scaled === true) {
        // §20.1.8.41: (cos x, sin x) becomes (w cos x, h sin x)
        // before normalization.
        dx *= tileW;
        dy *= tileH;
        const magnitude = Math.hypot(dx, dy);
        if (magnitude > 0) {
          dx /= magnitude;
          dy /= magnitude;
        }
      }
      const cx = tileX + tileW / 2;
      const cy = tileY + tileH / 2;
      const gradLen = (Math.abs(dx) * tileW + Math.abs(dy) * tileH) / 2;
      gradient = ctx.createLinearGradient(
        cx - dx * gradLen, cy - dy * gradLen,
        cx + dx * gradLen, cy + dy * gradLen,
      );
    }
    for (const stop of stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, stop.position)), hexToRgba(stop.color));
    }
    return gradient;
  }
  return null;
}

/**
 * Build a tiling CanvasPattern for an OOXML preset pattern fill.
 * Falls back to the foreground colour string when the preset name is unknown
 * or the OffscreenCanvas / Canvas environment cannot create a pattern.
 *
 * Cached per (preset, fg, bg) tuple — patterns are immutable bitmaps so the
 * same backing canvas can be reused across many shapes.
 */
const patternCache = new WeakMap<CanvasRenderingContext2D, Map<string, CanvasPattern>>();

function resolvePatternFill(
  fill: PatternFill,
  ctx: CanvasRenderingContext2D,
): CanvasPattern | string {
  const key = `${fill.preset}|${fill.fg}|${fill.bg}`;
  let perCtx = patternCache.get(ctx);
  if (!perCtx) {
    perCtx = new Map();
    patternCache.set(ctx, perCtx);
  }
  const cached = perCtx.get(key);
  if (cached) return cached;

  const bitmap = buildPatternBitmap(fill.preset, fill.fg, fill.bg);
  if (!bitmap) return hexToRgba(fill.fg);
  const pat = ctx.createPattern(bitmap, 'repeat');
  if (!pat) return hexToRgba(fill.fg);
  perCtx.set(key, pat);
  return pat;
}

/**
 * Apply a Stroke to ctx. `emuPerPx` converts stroke width from EMU to px
 * (e.g. scale factor from pptx's emuToPx).
 *
 * Parsers normalize symbolic dash vocabularies to DrawingML preset names; the
 * shared resolver also accepts VML's numeric relative grammar. Both scale by
 * the pixel line width and return `[]` for solid / unknown styles.
 */
export function applyStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke | null,
  emuPerPx: number,
): void {
  if (!stroke) {
    ctx.strokeStyle = 'transparent';
    ctx.lineWidth = 0;
    ctx.setLineDash([]);
    ctx.lineCap = 'butt';
    ctx.lineJoin = 'miter';
    ctx.miterLimit = 10;
    return;
  }
  ctx.strokeStyle = hexToRgba(stroke.color);
  const lw = Math.max(0.5, stroke.width * emuPerPx);
  ctx.lineWidth = lw;
  const dash = stroke.customDash != null
    ? drawingmlLineDashArray(stroke.customDash, null, lw)
    : stroke.dashStyle
      ? shapeStrokeDashArray(stroke.dashStyle, lw)
      : [];
  const requestedCap = stroke.lineCap ?? 'butt';
  // ECMA-376 Part 4 §19.1.2.21: a zero in VML's numeric dash grammar is a
  // visible fourfold-symmetric dot, even though VML's default endcap is flat.
  // Canvas drops a zero-length dash under its equivalent `butt` cap. A square
  // cap is the exact centered, fourfold-symmetric fallback for that one case;
  // explicit round/square caps keep their authored shape.
  const hasZeroDash = dash.some((length, index) => index % 2 === 0 && length === 0);
  ctx.lineCap = requestedCap === 'butt' && hasZeroDash ? 'square' : requestedCap;
  ctx.lineJoin = stroke.lineJoin ?? 'miter';
  ctx.miterLimit = stroke.miterLimit ?? 10;
  ctx.setLineDash(dash);
}
