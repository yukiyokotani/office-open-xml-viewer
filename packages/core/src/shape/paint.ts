import type { Fill, Stroke } from '../types/common';

/**
 * Convert a 6- or 8-char hex colour to a CSS `rgba()` string.
 * 8-char hex encodes alpha in the last two chars (RRGGBBAA).
 * `alpha` applies to 6-char hex; ignored for 8-char.
 */
export function hexToRgba(hex: string, alpha = 1): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const a = hex.length >= 8 ? parseInt(hex.slice(6, 8), 16) / 255 : alpha;
  return `rgba(${r},${g},${b},${a})`;
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
): string | CanvasGradient | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    const stops = fill.stops;
    if (stops.length === 0) return null;
    if (stops.length === 1) return hexToRgba(stops[0].color);

    let gradient: CanvasGradient;
    if (fill.gradType === 'radial') {
      const cx = x + w / 2;
      const cy = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    } else {
      const rad = (fill.angle * Math.PI) / 180;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const gradLen = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
      gradient = ctx.createLinearGradient(
        cx - Math.cos(rad) * gradLen, cy - Math.sin(rad) * gradLen,
        cx + Math.cos(rad) * gradLen, cy + Math.sin(rad) * gradLen,
      );
    }
    for (const stop of stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, stop.position)), hexToRgba(stop.color));
    }
    return gradient;
  }
  return null;
}

const DASH_PATTERNS: Record<string, number[]> = {
  dash:         [6, 3],
  dot:          [1.5, 3],
  dashDot:      [6, 3, 1.5, 3],
  lgDash:       [10, 4],
  lgDashDot:    [10, 4, 1.5, 4],
  lgDashDotDot: [10, 4, 1.5, 4, 1.5, 4],
  sysDash:      [4, 2],
  sysDot:       [1, 2],
  sysDashDot:   [4, 2, 1, 2],
};

/**
 * Apply a Stroke to ctx. `emuPerPx` converts stroke width from EMU to px
 * (e.g. scale factor from pptx's emuToPx).
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
    return;
  }
  ctx.strokeStyle = hexToRgba(stroke.color);
  const lw = Math.max(0.5, stroke.width * emuPerPx);
  ctx.lineWidth = lw;
  const pat = stroke.dashStyle ? DASH_PATTERNS[stroke.dashStyle] : null;
  ctx.setLineDash(pat ? pat.map((v) => v * lw) : []);
}
