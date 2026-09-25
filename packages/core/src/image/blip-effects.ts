// Shared CT_Blip pixel effects (ECMA-376 §20.1.8.13) beyond the duotone:
// `<a:grayscl>` (§20.1.8.34), `<a:biLevel>` (§20.1.8.11) and `<a:clrChange>`
// (§20.1.8.16), applied in document order to the decoded raster. The parser
// emits the list only when one of these effects is present; a `duotone` entry
// marks where the blip's `<a:duotone>` applies in that order.
//
// Office behaviour that the standard leaves implicit, and the evidence used:
// - Luminance. ECMA-376 converts to gray "corresponding to their luminance"
//   and thresholds bi-level by luminance without defining the weights.
//   PowerPoint's PDF export of grayscl + biLevel(50%) pictures turns
//   (2,167,223) white and (0,147,190), (0,126,229), (9,74,178) black; Rec. 601
//   weights put the first below 50% (0.486), Rec. 709 weights (0.532) and the
//   others on the observed sides, so the effects use Rec. 709 luma of the
//   stored (gamma-encoded) values.
// - clrChange. MS-OI29500 says Office leaves alpha alone unless useA is set,
//   but PowerPoint's own "set transparent colour" output (clrTo = the same
//   colour with alpha 0, useA absent) renders transparent in its PDF export.
//   Exact RGB matches therefore take clrTo's colour and alpha; with useA the
//   source alpha must match clrFrom's alpha as well.
// - lum. ECMA-376 §20.1.8.42 names brightness and contrast but gives no
//   formula. Evidence: PowerPoint's PDF export of a 256-step gray ramp under
//   a grid of bright × contrast values (±35%, ±70%, ±100%, and 0). Every
//   channel follows
//     out = clamp(k · (v + b/2 − ½) + ½ + b/2),
//   with v and out in 0–1 and b the brightness fraction. k = 1 + c for
//   contrast c ≤ 0 and 1 / (1 − c) for c > 0. So half of the brightness
//   shift is applied before the contrast scaling around mid-gray and half
//   after. This matches every measured level to within 0.7 of 255. The ramp
//   saved as a binary .ppt matches it too, with the stored contrast used
//   directly as k. At c = 1 the result is a threshold at v + b/2 = ½.

import type { RgbaBuffer, Duotone } from './duotone';
import { duotoneImageData, hex6ToRgb } from './duotone';

/** One CT_Blip pixel effect, as emitted by the parsers (camelCase JSON). */
export type BlipEffect =
  | { type: 'grayscale' }
  | { type: 'biLevel'; thresh: number }
  | {
      type: 'colorChange';
      from: string;
      fromAlpha: number;
      to: string;
      toAlpha: number;
      useAlpha: boolean;
    }
  | { type: 'luminance'; bright: number; contrast: number }
  | { type: 'duotone' };

/** A picture's pixel transform: ordered effects plus the duotone colours the
 *  `duotone` entry refers to. */
export interface BlipPixelEffects {
  readonly effects: readonly BlipEffect[];
  readonly duotone?: Duotone | null;
}

export function isBlipPixelEffects(value: unknown): value is BlipPixelEffects {
  return typeof value === 'object' && value !== null && Array.isArray((value as BlipPixelEffects).effects);
}

/** Stable cache-key suffix for a pixel transform. */
export function blipPixelEffectsKey(value: BlipPixelEffects): string {
  const parts = value.effects.map((effect) => {
    switch (effect.type) {
      case 'grayscale':
        return 'g';
      case 'biLevel':
        return `b${effect.thresh}`;
      case 'colorChange':
        return `c${effect.from}.${effect.fromAlpha}>${effect.to}.${effect.toAlpha}${effect.useAlpha ? 'a' : ''}`;
      case 'luminance':
        return `l${effect.bright}.${effect.contrast}`;
      case 'duotone':
        return value.duotone ? `d${value.duotone.clr1}.${value.duotone.clr2}` : 'd';
    }
  });
  // A duotone without a position marker (a producer that placed it directly
  // under blipFill) applies after the listed effects.
  if (value.duotone && !value.effects.some((effect) => effect.type === 'duotone')) {
    parts.push(`d${value.duotone.clr1}.${value.duotone.clr2}`);
  }
  return parts.join(',');
}

/** Rec. 709 luma (0–1) of gamma-encoded sRGB bytes; see the evidence above. */
export function blipLuminance(r: number, g: number, b: number): number {
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
}

/** Apply the effects in place, in order. Alpha is preserved except by
 *  clrChange, whose target alpha replaces the matched pixel's. */
export function applyBlipPixelEffects(buf: RgbaBuffer, value: BlipPixelEffects): RgbaBuffer {
  const d = buf.data;
  let duotoneApplied = false;
  for (const effect of value.effects) {
    switch (effect.type) {
      case 'grayscale':
        for (let i = 0; i < d.length; i += 4) {
          const gray = Math.round(blipLuminance(d[i], d[i + 1], d[i + 2]) * 255);
          d[i] = gray;
          d[i + 1] = gray;
          d[i + 2] = gray;
        }
        break;
      case 'biLevel':
        for (let i = 0; i < d.length; i += 4) {
          // "Values greater than or equal to the threshold are set to white."
          const level = blipLuminance(d[i], d[i + 1], d[i + 2]) >= effect.thresh ? 255 : 0;
          d[i] = level;
          d[i + 1] = level;
          d[i + 2] = level;
        }
        break;
      case 'colorChange': {
        const from = hex6ToRgb(effect.from);
        const to = hex6ToRgb(effect.to);
        if (!from || !to) break;
        const fromAlpha = Math.round(effect.fromAlpha * 255);
        const toAlpha = Math.round(effect.toAlpha * 255);
        for (let i = 0; i < d.length; i += 4) {
          if (d[i] !== from[0] || d[i + 1] !== from[1] || d[i + 2] !== from[2]) continue;
          if (effect.useAlpha && d[i + 3] !== fromAlpha) continue;
          d[i] = to[0];
          d[i + 1] = to[1];
          d[i + 2] = to[2];
          d[i + 3] = toAlpha;
        }
        break;
      }
      case 'luminance': {
        const table = luminanceTable(effect.bright, effect.contrast);
        for (let i = 0; i < d.length; i += 4) {
          d[i] = table[d[i]];
          d[i + 1] = table[d[i + 1]];
          d[i + 2] = table[d[i + 2]];
        }
        break;
      }
      case 'duotone':
        if (value.duotone) duotoneImageData(buf, value.duotone.clr1, value.duotone.clr2);
        duotoneApplied = true;
        break;
    }
  }
  if (!duotoneApplied && value.duotone) duotoneImageData(buf, value.duotone.clr1, value.duotone.clr2);
  return buf;
}

/** Per-channel lookup table for `<a:lum>` (see the evidence above): bright
 *  and contrast are signed fractions in [-1, 1]. */
export function luminanceTable(bright: number, contrast: number): Uint8ClampedArray {
  const table = new Uint8ClampedArray(256);
  const b = Math.max(-1, Math.min(1, bright));
  const c = Math.max(-1, Math.min(1, contrast));
  for (let v = 0; v < 256; v++) {
    const shifted = v / 255 + b / 2 - 0.5;
    let out: number;
    if (c >= 1) {
      out = shifted > 0 ? 1 : shifted < 0 ? 0 : 0.5 + b / 2;
    } else {
      const k = c <= 0 ? 1 + c : 1 / (1 - c);
      out = k * shifted + 0.5 + b / 2;
    }
    table[v] = Math.round(Math.max(0, Math.min(1, out)) * 255);
  }
  return table;
}
