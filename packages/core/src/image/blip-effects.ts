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
//   others on the observed sides. Boundary evidence: a layout picture that
//   PowerPoint 16 reads from a binary .ppt as grayscl + biLevel(50%), compared
//   with PowerPoint's PDF of that .ppt at every PDF pixel whose source block is
//   a single colour that candidate rules classify differently (7,068 blocks of
//   JPEG blues around the threshold): Rec. 709 luma of the stored
//   (gamma-encoded) values with the grayscale truncated to a whole 8-bit level
//   matches all 7,068; rounding that level instead fails 30 (e.g.
//   (99,131,178): luma 127.59 must become 127, black), unquantized Rec. 709
//   fails 30, Rec. 601 fails at least 1,534, the channel average 4,553 and
//   linear-light Rec. 709 is contradicted on nearly every block. So grayscl
//   writes floor(Rec. 709 luma) and biLevel compares Rec. 709 luma with the
//   threshold. (No block sat exactly on a threshold, so >= follows the text:
//   "values greater than or equal to the threshold are set to white".)
// - clrChange. MS-OI29500 says Office leaves alpha alone unless useA is set,
//   but PowerPoint's own "set transparent colour" output (clrTo = the same
//   colour with alpha 0, useA absent) renders transparent in its PDF export:
//   comparing the PDF soft masks with the source pictures, every block that is
//   pure clrFrom white becomes transparent (2,286 blocks with no counterexample
//   in one picture, 291 in another; 971 of them transparent where keeping the
//   source alpha would leave them opaque), and every non-matching block stays
//   opaque. Exact RGB matches therefore take clrTo's colour and alpha; with
//   useA the source alpha must match clrFrom's alpha as well.
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
import {
  MAX_IMAGE_EFFECT_PASSES,
  MAX_IMAGE_EFFECT_PIXEL_WORK,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';

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

/** grayscl's 8-bit level: the Rec. 709 luma truncated to a whole level (see
 *  the evidence above). The epsilon keeps white at 255 despite the weights'
 *  binary rounding. */
export function blipGrayLevel(r: number, g: number, b: number): number {
  return Math.floor(blipLuminance(r, g, b) * 255 + 1e-6);
}

/** Full pixel passes a transform runs: one per listed effect, plus the
 *  duotone when it applies after the list (no position marker). */
export function blipPixelEffectPasses(value: BlipPixelEffects): number {
  const trailingDuotone = value.duotone
    && !value.effects.some((effect) => effect.type === 'duotone');
  return value.effects.length + (trailingDuotone ? 1 : 0);
}

/**
 * Reject a transform whose pass count or cumulative pixel work exceeds the
 * shared resource policy (`MAX_IMAGE_EFFECT_PASSES`,
 * `MAX_IMAGE_EFFECT_PIXEL_WORK` in pixel-budget.ts) before any pixel is
 * touched. A crossing is an `OoxmlDecodedImageLimitError`, the same catchable
 * quota failure as an over-budget decode; effects are never dropped or
 * truncated to fit. `pixels` 0 checks the pass count alone (before decoding).
 */
export function assertBlipPixelEffectsBudget(value: BlipPixelEffects, pixels: number): void {
  const passes = blipPixelEffectPasses(value);
  if (passes > MAX_IMAGE_EFFECT_PASSES) {
    throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, passes);
  }
  const work = passes * Math.max(0, Math.ceil(pixels));
  if (!Number.isSafeInteger(work) || work > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError(
      'image-effect-work',
      MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(work) ? work : Number.MAX_SAFE_INTEGER,
    );
  }
}

/** Apply the effects in place, in order. Alpha is preserved except by
 *  clrChange, whose target alpha replaces the matched pixel's. Throws
 *  `OoxmlDecodedImageLimitError` before touching the buffer when the transform
 *  exceeds the effect budget (see {@link assertBlipPixelEffectsBudget}). */
export function applyBlipPixelEffects(buf: RgbaBuffer, value: BlipPixelEffects): RgbaBuffer {
  const d = buf.data;
  assertBlipPixelEffectsBudget(value, d.length / 4);
  let duotoneApplied = false;
  for (const effect of value.effects) {
    switch (effect.type) {
      case 'grayscale':
        for (let i = 0; i < d.length; i += 4) {
          const gray = blipGrayLevel(d[i], d[i + 1], d[i + 2]);
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
