import { describe, it, expect } from 'vitest';
import { measureAdvanceWithTrailingSpace } from './measure-space';

// Model of a Canvas 2D context whose measureText advance is a fixed per-glyph
// sum, with a switch for the trailing-whitespace divergence.
//
//   - CHAR_W per non-space glyph, SPACE_W per space.
//   - `trims=true` models Firefox/WebKit/Safari: TRAILING spaces (and a wholly
//     whitespace string) contribute 0; interior spaces still count.
//   - `trims=false` models Chrome/Blink & skia-canvas: every space counts.
const CHAR_W = 10;
const SPACE_W = 4;

function makeCtx(trims: boolean, font = '16px serif'): {
  ctx: CanvasRenderingContext2D;
  fontRef: { value: string };
} {
  const fontRef = { value: font };
  const measure = (s: string): number => {
    const eff = trims ? s.replace(/ +$/, '') : s;
    let w = 0;
    for (const ch of eff) w += ch === ' ' ? SPACE_W : CHAR_W;
    return w;
  };
  const ctx = {
    get font() { return fontRef.value; },
    set font(v: string) { fontRef.value = v; },
    measureText: (s: string) => ({ width: measure(s) }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fontRef };
}

describe('measureAdvanceWithTrailingSpace', () => {
  it('non-trimming engine: identical to raw measureText (pass-through)', () => {
    const { ctx } = makeCtx(false);
    // "Other " = 5×10 + 4 = 54, "countries" = 9×10 = 90.
    expect(measureAdvanceWithTrailingSpace(ctx, 'Other ')).toBe(54);
    expect(measureAdvanceWithTrailingSpace(ctx, 'countries')).toBe(90);
  });

  it('trimming engine: restores the trailing-space advance', () => {
    const { ctx } = makeCtx(true);
    // Raw would trim "Other " → "Other" = 50; the helper adds one SPACE_W back.
    expect(measureAdvanceWithTrailingSpace(ctx, 'Other ')).toBe(54);
    // No trailing space ⇒ unchanged.
    expect(measureAdvanceWithTrailingSpace(ctx, 'countries')).toBe(90);
  });

  it('trimming engine: a lone space token (pptx) keeps its advance', () => {
    const { ctx } = makeCtx(true);
    // Raw measureText(' ') trims to '' → 0; the helper restores SPACE_W.
    expect(measureAdvanceWithTrailingSpace(ctx, ' ')).toBe(SPACE_W);
    expect(measureAdvanceWithTrailingSpace(ctx, '   ')).toBe(3 * SPACE_W);
  });

  it('interior spaces are never affected on either engine', () => {
    for (const trims of [false, true]) {
      const { ctx } = makeCtx(trims);
      // "a b" = 10 + 4 + 10 = 24 (interior space always counts).
      expect(measureAdvanceWithTrailingSpace(ctx, 'a b')).toBe(24);
    }
  });

  it('empty string is 0', () => {
    expect(measureAdvanceWithTrailingSpace(makeCtx(true).ctx, '')).toBe(0);
    expect(measureAdvanceWithTrailingSpace(makeCtx(false).ctx, '')).toBe(0);
  });

  it('detection is per-font (a size change re-probes and still corrects)', () => {
    const { ctx, fontRef } = makeCtx(true);
    expect(measureAdvanceWithTrailingSpace(ctx, 'x ')).toBe(CHAR_W + SPACE_W);
    fontRef.value = '32px serif';
    // Same stub advances (font size does not scale the stub), but a fresh font
    // string forces a re-probe; the correction must still apply.
    expect(measureAdvanceWithTrailingSpace(ctx, 'x ')).toBe(CHAR_W + SPACE_W);
  });
});
