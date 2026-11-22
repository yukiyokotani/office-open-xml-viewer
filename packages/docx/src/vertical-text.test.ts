import { describe, it, expect } from 'vitest';
import {
  isUprightVerticalGlyph,
  verticalDrawMode,
  splitVerticalOrientationRuns,
  drawVerticalRun,
  drawVerticalRunWithCapability,
  drawTateChuYokoRun,
  drawUprightBox,
  planVerticalRunWithCapability,
  verticalRunInkExtraPx,
  verticalRunInkExtraPxWithCapability,
} from './vertical-text.js';

// ECMA-376 §17.6.20 vertical writing (tbRl). These are the pure classification
// + geometry primitives the renderer wires into the glyph/image draw path behind
// the `verticalCJK` flag. Orientation is decided by the Unicode UAX#50
// Vertical_Orientation property (core `verticalOrientation`).

describe('verticalDrawMode (UAX#50 vo → draw mode)', () => {
  const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

  it('U/Tu → upright (ideographs, kana, 、。！？, small kana)', () => {
    for (const ch of ['富', '士', 'あ', 'ア', '、', '。', '！', '？', 'ぁ', 'ッ']) {
      expect(verticalDrawMode(cp(ch))).toBe('upright');
    }
  });

  it('Tr → rotate (long vowel mark ー, corner brackets, parens, quotes)', () => {
    for (const ch of ['ー', '「', '」', '（', '）', '〈', '〉', '“', '”']) {
      expect(verticalDrawMode(cp(ch))).toBe('rotate');
    }
  });

  it('R → sideways (Latin, digits, ASCII punctuation)', () => {
    for (const ch of ['A', 'z', '0', '5', '9', '@', '-', '.']) {
      expect(verticalDrawMode(cp(ch))).toBe('sideways');
    }
  });
});

describe('planVerticalRunWithCapability (retained vertical paint geometry)', () => {
  it.each([2, -1])('keeps vertical retained origins aligned with mark-free tracking at %s pt', pitch => {
    const { ctx } = mockCtx();
    ctx.measureText = (text: string) => ({ width: [...text].filter(c => c === 'a' || c === 'b' || c === '字').length * 10 } as TextMetrics);
    const cells = planVerticalRunWithCapability(ctx, 'a\u030Ab字', 10, pitch);
    expect(cells.map(c => [c.text, c.originPt, c.advancePt])).toEqual([
      ['a\u030Ab', 0, 20 + 2 * pitch], ['字', 20 + 2 * pitch + (10 + pitch) / 2, 10 + pitch],
    ]);
  });

  // #1975 WORD_CONDENSED_SPACING_ADVANCE_FLOOR: with acquisition's opt-in an
  // over-condensed character advances zero; a sideways piece that floors
  // paints per cluster. Without it the plan is the ordinary unfloored one.
  it('floors over-condensed vertical cells at zero advance', () => {
    const floor = () => 0;
    const { ctx } = mockCtx();
    ctx.measureText = (text: string) => ({
      width: [...text].reduce((sum, c) => sum + (c === 'i' ? 4 : 10), 0),
    } as TextMetrics);
    const cells = planVerticalRunWithCapability(ctx, 'Mi字字', 10, -6, 1, false, () => false, 'vertical-rl', 0, floor);
    expect(cells.map(c => [c.text, c.orientation, c.advancePt])).toEqual([
      ['M', 'sideways', 4], ['i', 'sideways', 0], ['字', 'upright', 4], ['字', 'upright', 4],
    ]);
    expect(cells.map(c => c.originPt)).toEqual([0, 4, 6, 10]);
    const deep = planVerticalRunWithCapability(ctx, '字', 10, -30, 1, false, () => false, 'vertical-rl', 0, floor);
    expect(deep[0]!.advancePt).toBe(0);
    // Unfloored pieces keep one contextual operation.
    expect(planVerticalRunWithCapability(ctx, 'MM', 10, -2, 1, false, () => false, 'vertical-rl', 0, floor)
      .map(c => c.text)).toEqual(['MM']);
    // No opt-in (e.g. fitText's own allocation): the ordinary signed plan.
    expect(planVerticalRunWithCapability(ctx, '字', 10, -30)[0]!.advancePt).toBe(-20);
  });

  // The floor covers each character's punctuation compression / width balance
  // supplied by acquisition, whatever their sign.
  it('floors vertical cells together with their character adjustments', () => {
    const { ctx } = mockCtx();
    ctx.measureText = (text: string) => ({
      width: [...text].reduce((sum, c) => sum + (c === 'M' ? 20 : c === ' ' ? 12 : 10), 0),
    } as TextMetrics);
    const positive = planVerticalRunWithCapability(
      ctx, '  ', 10, -30, 1, false, () => false, 'vertical-rl', 0, () => 5,
    );
    expect(positive.map(c => [c.advancePt, c.floored ?? false])).toEqual([[0, true], [0, true]]);
    const balanced = planVerticalRunWithCapability(
      ctx, 'M  ', 10, -11, 1, false, () => false, 'vertical-rl', 0,
      (start, end) => (start === 0 && end === 1 ? 0 : -2),
    );
    expect(balanced.map(c => [c.text, c.advancePt])).toEqual([['M', 9], [' ', 0], [' ', 0]]);
  });

  it('keeps a floored cursive sideways piece as one cell advancing by its floored sum', () => {
    const { ctx } = mockCtx();
    ctx.measureText = (text: string) => ({
      width: [...text].reduce((sum, c) => sum + (c === 'ي' ? 4 : 10), 0),
    } as TextMetrics);
    const cells = planVerticalRunWithCapability(
      ctx, 'بيت', 10, -6, 1, false, () => false, 'vertical-rl', 0, () => 0,
    );
    expect(cells.map(c => [c.text, c.advancePt, c.floored ?? false])).toEqual([['بيت', 8, true]]);
  });

  it('retains the font em-box cross-axis correction for a sideways phone number', () => {
    const { ctx } = mockCtx({
      fontBoundingBoxAscent: 10,
      fontBoundingBoxDescent: 2,
    });

    const cells = planVerticalRunWithCapability(
      ctx,
      '電話03-1234-5678',
      12,
      0,
      1,
      false,
      () => false,
    );

    expect(cells.map((cell) => ({
      text: cell.text,
      orientation: cell.orientation,
      crossAxisOffsetPt: cell.drawOffsetPt.yPt,
    }))).toEqual([
      { text: '電', orientation: 'upright', crossAxisOffsetPt: 0 },
      { text: '話', orientation: 'upright', crossAxisOffsetPt: 0 },
      // Sideways text uses an alphabetic baseline. Retain the font-level
      // (ascent - descent) / 2 correction so paint can share the same physical
      // column centre as the upright cells without measuring the font again.
      { text: '03-1234-5678', orientation: 'sideways', crossAxisOffsetPt: 4 },
    ]);
  });

  it('retains UAX #50 Tr substitution and rotate fallbacks with original source ranges', () => {
    const { ctx } = mockCtx();

    const cells = planVerticalRunWithCapability(ctx, 'A（ー）B', 12, 0, 1, true, () => false);

    expect(cells.map(({ text, orientation, range, verticalFeature }) => ({
      text, orientation, range, verticalFeature,
    }))).toEqual([
      { text: 'A', orientation: 'sideways', range: { start: 0, end: 1 }, verticalFeature: false },
      { text: '︵', orientation: 'upright', range: { start: 1, end: 2 }, verticalFeature: false },
      { text: 'ー', orientation: 'rotate', range: { start: 2, end: 3 }, verticalFeature: false },
      { text: '︶', orientation: 'upright', range: { start: 3, end: 4 }, verticalFeature: false },
      { text: 'B', orientation: 'sideways', range: { start: 4, end: 5 }, verticalFeature: false },
    ]);
    // Unicode presentation-form fallbacks are ordinary centred Canvas glyphs.
    // Feature-backed placement adjustments must not leak into this route.
    expect(cells.map(({ drawOffsetPt }) => drawOffsetPt.xPt)).toEqual([0, 0, 0, 0, 0]);
  });

  it('retains original transform characters and projects their vertical origin into the cell', () => {
    const { ctx } = mockCtx({
      vert: {
        '（': { width: 12, asc: 9, desc: 3 },
        'ー': { width: 12, asc: 9, desc: 3 },
        '）': { width: 12, asc: 9, desc: 3 },
      },
    });

    const cells = planVerticalRunWithCapability(ctx, '（ー）', 12, 0, 1, true, () => true);

    expect(cells.map(({ text, orientation, verticalFeature }) => ({
      text, orientation, verticalFeature,
    }))).toEqual([
      { text: '（', orientation: 'upright', verticalFeature: true },
      { text: 'ー', orientation: 'upright', verticalFeature: true },
      { text: '）', orientation: 'upright', verticalFeature: true },
    ]);
    expect(cells.map(({ drawOffsetPt }) => drawOffsetPt.xPt)).toEqual([0, 0, 0]);
  });

  it('retains counter-rotated logical block ink for the upright base glyph', () => {
    const { ctx } = mockCtx({
      inkLR: { '．': { left: 2, right: 6 } },
    });

    const [cell] = planVerticalRunWithCapability(
      ctx,
      '．',
      10,
      0,
      0.5,
      true,
      () => false,
    );

    expect(cell).toMatchObject({
      orientation: 'upright',
      drawOffsetPt: { xPt: 0, yPt: 0 },
      // Paint rotates the local x ink [-2, 6] onto logical -y.
      // vertical-rl w:w scales local y, so the 0.5 scale does not alter this
      // block-axis interval.
      blockAxisInkBounds: { startPt: -6, endPt: 2 },
    });
  });
});

describe('isUprightVerticalGlyph (UAX#50 vo ∈ {U, Tu})', () => {
  const cp = (ch: string): number => ch.codePointAt(0) ?? 0;

  it('is true for U/Tu (ideographs, kana, 、。) and false for Tr/R', () => {
    for (const ch of ['富', 'あ', '、', '。']) expect(isUprightVerticalGlyph(cp(ch))).toBe(true);
    // ー「」（） are Tr (rotate), Latin/digits are R — not upright.
    for (const ch of ['ー', '「', '）', 'A', '5']) expect(isUprightVerticalGlyph(cp(ch))).toBe(false);
  });
});

describe('splitVerticalOrientationRuns (§17.6.20 — group by draw mode)', () => {
  it('splits a mixed run into maximal same-mode pieces in logical order', () => {
    const pieces = splitVerticalOrientationRuns('第5回大会');
    expect(pieces).toEqual([
      { text: '第', mode: 'upright' },
      { text: '5', mode: 'sideways' },
      { text: '回大会', mode: 'upright' },
    ]);
  });

  it('separates a Tr bracket/長音符 into its own rotate piece', () => {
    // チーム(土): チ=upright, ー=rotate(Tr), ム=upright, (=rotate(Tr), 土=upright, )=rotate(Tr)
    const pieces = splitVerticalOrientationRuns('チーム（土）');
    expect(pieces).toEqual([
      { text: 'チ', mode: 'upright' },
      { text: 'ー', mode: 'rotate' },
      { text: 'ム', mode: 'upright' },
      { text: '（', mode: 'rotate' },
      { text: '土', mode: 'upright' },
      { text: '）', mode: 'rotate' },
    ]);
  });

  it('keeps a pure-CJK run as one upright piece', () => {
    expect(splitVerticalOrientationRuns('富士町')).toEqual([{ text: '富士町', mode: 'upright' }]);
  });

  it('keeps a pure-Latin run as one sideways piece', () => {
    expect(splitVerticalOrientationRuns('2026')).toEqual([{ text: '2026', mode: 'sideways' }]);
  });

  it('returns nothing for empty text', () => {
    expect(splitVerticalOrientationRuns('')).toEqual([]);
  });

  it('preserves surrogate pairs as single code points', () => {
    const pieces = splitVerticalOrientationRuns('𠀋'); // CJK Ext-B ideograph (surrogate pair)
    expect(pieces).toHaveLength(1);
    expect(pieces[0].text).toBe('𠀋');
  });
});

// A minimal 2D-context spy recording the transform + text/box draw ops so we can
// assert the draw geometry without a real canvas.
type Op =
  | { op: 'save' }
  | { op: 'restore' }
  | { op: 'translate'; x: number; y: number }
  | { op: 'rotate'; a: number }
  | { op: 'scale'; sx: number; sy: number }
  | { op: 'transform'; a: number; b: number; c: number; d: number; e: number; f: number }
  | { op: 'fillText'; text: string; x: number; y: number; align: string; baseline: string; feature: string }
  | { op: 'draw'; dx: number; dy: number; dw: number; dh: number };

// Optional metrics the mock returns from `measureText`, keyed by the metric it
// is asked to model. `fontBoundingBox*` are font-level (glyph-independent, used
// for the sideways em-box-centre shift); `actualBoundingBox*` are per-glyph tight
// ink bounds (used for the upright along-column ink centring). Values are read at
// the CURRENT textBaseline the helper sets before measuring, so the mock reports
// the same numbers regardless — the helper does the (asc−desc)/2 arithmetic.
interface MockMetrics {
  fontBoundingBoxAscent?: number;
  fontBoundingBoxDescent?: number;
  // Per-glyph tight ink extent under a `middle` textBaseline, by draw glyph.
  inkMiddle?: Record<string, { asc: number; desc: number }>;
  // Per-glyph HORIZONTAL tight ink bounds relative to the advance CENTRE (the
  // values a `textAlign='center'` measureText reports as
  // actualBoundingBoxLeft/Right), by glyph. Used by the vo=Tr rotate-fallback
  // ink-overrun path (#1014): after the +90° page rotation the glyph's HORIZONTAL
  // ink maps onto the along-column axis, so left+right is the along-column ink
  // extent. Returned regardless of the current textAlign (the values are already
  // centre-relative).
  inkLR?: Record<string, { left: number; right: number }>;
  // Metrics returned only while the composed OpenType `vert` feature is active.
  // They model the original code point's feature-selected glyph and placement.
  vert?: Record<string, { width?: number; asc: number; desc: number }>;
  // Whole-run widths can include horizontal kern-pair compression even though
  // vertical paint advances one independent glyph cell at a time.
  wholeWidths?: Record<string, number>;
  shearSlope?: number;
}

function mockCtx(metrics: MockMetrics = {}): { ctx: any; ops: Op[] } {
  const ops: Op[] = [];
  const style = { fontFeatureSettings: 'normal' };
  class ScratchCanvas {
    width: number;
    height: number;
    style = style;
    constructor(width: number, height: number) { this.width = width; this.height = height; }
    getContext() {
      const canvas = this;
      return {
        canvas,
        font: '', fillStyle: '#000', textAlign: 'center', textBaseline: 'middle',
        clearRect() {}, fillText() {},
        getImageData() {
          const data = new Uint8ClampedArray(canvas.width * canvas.height * 4);
          const slope = metrics.shearSlope ?? 0;
          for (let x = 128; x <= 384; x++) {
            const y = Math.round(256 + slope * (x - 256));
            data[(y * canvas.width + x) * 4 + 3] = 255;
          }
          return { data };
        },
      };
    }
  }
  const ctx: any = {
    canvas: metrics.shearSlope === undefined ? { style } : new ScratchCanvas(1, 1),
    font: '12px serif',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    letterSpacing: '0px',
    save() {
      ops.push({ op: 'save' });
    },
    restore() {
      ops.push({ op: 'restore' });
    },
    translate(x: number, y: number) {
      ops.push({ op: 'translate', x, y });
    },
    rotate(a: number) {
      ops.push({ op: 'rotate', a });
    },
    scale(sx: number, sy: number) {
      ops.push({ op: 'scale', sx, sy });
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      ops.push({ op: 'transform', a, b, c, d, e, f });
    },
    measureText(s: string) {
      // Every code point is 10px wide.
      const vert = style.fontFeatureSettings.includes('"vert" 1') ? metrics.vert?.[s] : undefined;
      const m: Record<string, number> = {
        width: vert?.width ?? metrics.wholeWidths?.[s] ?? [...s].length * 10,
      };
      if (metrics.fontBoundingBoxAscent !== undefined) {
        m.fontBoundingBoxAscent = metrics.fontBoundingBoxAscent;
        m.fontBoundingBoxDescent = metrics.fontBoundingBoxDescent ?? 0;
      }
      const ink = vert ?? metrics.inkMiddle?.[s];
      if (ink && this.textBaseline === 'middle') {
        m.actualBoundingBoxAscent = ink.asc;
        m.actualBoundingBoxDescent = ink.desc;
      }
      const lr = metrics.inkLR?.[s];
      if (lr) {
        m.actualBoundingBoxLeft = lr.left;
        m.actualBoundingBoxRight = lr.right;
      }
      return m;
    },
    fillText(text: string, x: number, y: number) {
      ops.push({
        op: 'fillText', text, x, y, align: this.textAlign,
        baseline: this.textBaseline, feature: style.fontFeatureSettings,
      });
    },
  };
  return { ctx, ops };
}

describe('drawVerticalRun (§17.6.20 — upright CJK counter-rotated, Latin sideways)', () => {
  it.each([
    { fontPx: 12, ink: undefined, expectedY: 0 },
    { fontPx: 24, ink: { asc: 2, desc: 10 }, expectedY: -4 },
  ])('uses the ordinary upright fallback for an unreachable fullwidth period at $fontPx px', ({ fontPx, ink, expectedY }) => {
    const { ctx, ops } = mockCtx({ inkMiddle: ink ? { '．': ink } : undefined });
    const cells = planVerticalRunWithCapability(ctx, '．', fontPx, 0, 1, false, () => false);
    drawVerticalRunWithCapability(ctx, '．', 0, 0, fontPx, 0, 1, false, () => false);
    const fills = ops.filter((op): op is Extract<Op, { op: 'fillText' }> => op.op === 'fillText');
    // UAX #50 Tu permits the original upright glyph when its vertical alternate
    // is inaccessible. Horizontal ink metrics cannot recover that alternate.
    expect(fills).toEqual([
      { op: 'fillText', text: '．', x: 0, y: expectedY,
        align: 'center', baseline: 'middle', feature: 'normal' },
    ]);
    expect(cells[0].drawOffsetPt).toEqual({ xPt: 0, yPt: expectedY });
    expect(cells[0].orientation).toBe('upright');
    expect(cells[0].verticalFeature).toBe(false);
  });

  it('uses the original code point under vert for every reachable Tu/Tr glyph', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRunWithCapability(
      ctx,
      'ー〜～、。：；「」“”A',
      0,
      0,
      12,
      0,
      1,
      true,
      () => true,
    );
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => fill.text)).toEqual([
      'ー', '〜', '～', '、', '。', '：', '；', '「', '」', '“', '”', 'A',
    ]);
    expect(fills.map((fill) => fill.feature)).toEqual([
      ...Array.from({ length: 11 }, () => '"vert" 1'),
      'normal',
    ]);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(11);
    expect(rotates.every((rotate) => rotate.a === -Math.PI / 2)).toBe(true);
    expect(ops.some((op) => op.op === 'scale' && op.sy === -1)).toBe(false);
  });

  it('keeps every manual FE/upright fallback when vert is unreachable', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRunWithCapability(
      ctx,
      'ー〜～、。：；「」“”A',
      0,
      0,
      12,
      0,
      1,
      true,
      () => false,
    );
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => fill.text)).toEqual([
      'ー', '〜', '～', '︑', '︒', '：', '；', '﹁', '﹂', '“', '”', 'A',
    ]);
    expect(fills.every((fill) => fill.feature === 'normal')).toBe(true);
  });

  it('counter-rotates every upright glyph −90° about its cell centre', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '富士', 100, 200, 12, 0);
    // Two upright glyphs → two save/rotate(−90°)/restore triples.
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(rotates.every((r) => Math.abs(r.a - -Math.PI / 2) < 1e-9)).toBe(true);
    // First cell centre: x=100 + adv/2 (adv=10) = 105, baseline y=200.
    const firstTranslate = ops.find((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(firstTranslate).toEqual({ op: 'translate', x: 105, y: 200 });
    // Upright glyphs draw centred.
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.every((f) => f.align === 'center' && f.baseline === 'middle')).toBe(true);
  });

  it('draws a Latin glyph sideways (no rotation, alphabetic baseline, at the advance x)', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'A', 100, 200, 12, 0);
    expect(ops.some((o) => o.op === 'rotate')).toBe(false);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // Sideways: left at run x (advance 0), alphabetic baseline. The cross-axis y
    // is shifted down by the em-box centre so the glyph's ink centres on the same
    // column line the upright cells use. With no fontBoundingBox metrics the mock
    // falls back to 0.38·fontPx = 4.56, so y = 200 + 4.56.
    expect(fill?.text).toBe('A');
    expect(fill?.x).toBe(100);
    expect(fill?.baseline).toBe('alphabetic');
    expect(fill?.y).toBeCloseTo(200 + 0.38 * 12, 6);
  });

  it('centres a sideways glyph on the em-box centre from fontBoundingBox metrics (§17.6.20)', () => {
    // Symptom 1: mixed columns like "電話 03-…" must share one centreline. The
    // sideways glyph is drawn on its alphabetic baseline shifted down by the
    // font's em-box centre = (fontBoundingBoxAscent − fontBoundingBoxDescent)/2.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 40, fontBoundingBoxDescent: 10 });
    drawVerticalRun(ctx, '0', 100, 200, 48, 0);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // emBoxCenter = (40 − 10)/2 = 15 → y = 200 + 15 = 215.
    expect(fill).toMatchObject({ text: '0', x: 100, baseline: 'alphabetic' });
    expect(fill?.y).toBeCloseTo(215, 6);
  });

  it('advances a shaped sideways piece by measure + per-glyph letterSpacing', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'AB富', 0, 0, 12, 4);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // AB is one 20px shaped piece plus two 4px pitches. The following 10px CJK
    // cell therefore centres at x=28 + (10 + 4)/2 = 35.
    expect(fills.map((f) => [f.text, f.x])).toEqual([['AB', 0], ['富', 0]]);
    const translate = ops.find(
      (op): op is Extract<Op, { op: 'translate' }> => op.op === 'translate',
    );
    expect(translate).toEqual({ op: 'translate', x: 35, y: 0 });
  });

  it('uses the plain page-frame rotation for an unreachable Tr long mark', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, 'ー', 100, 200, 12, 0);
    // The enclosing tbRl page transform supplies +90°. The unreachable fallback
    // adds no counter-rotation, reflection, or shear.
    expect(ops.some((o) => o.op === 'rotate')).toBe(false);
    expect(ops.some((o) => o.op === 'translate')).toBe(false);
    expect(ops.some((o) => o.op === 'transform')).toBe(false);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill).toMatchObject({ text: 'ー', x: 105, y: 200, align: 'center', baseline: 'middle' });
  });

  it('uses the same plain rotation fallback for wave dash and fullwidth tilde', () => {
    for (const ch of ['〜', '～']) {
      const { ctx, ops } = mockCtx();
      drawVerticalRun(ctx, ch, 0, 0, 12, 0);
      expect(ops.some((o) => o.op === 'rotate')).toBe(false);
      expect(ops.some((o) => o.op === 'transform')).toBe(false);
      const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
      expect(fill).toMatchObject({ text: ch, x: 5, y: 0, align: 'center', baseline: 'middle' });
    }
  });

  it('does NOT reflect a Tr rotate glyph whose vertical form is a pure rotation (quotes “ ”)', () => {
    // The double quotes are vo=Tr rotate-fallback, but their designed vertical form IS
    // the +90° rotation (font-verified), so they must NOT reflect — they keep the plain
    // fillText in the page frame at the cell centre, with NO scale and NO translate.
    for (const ch of ['“', '”']) {
      const { ctx, ops } = mockCtx();
      drawVerticalRun(ctx, ch, 100, 200, 12, 0);
      expect(ops.some((o) => o.op === 'rotate')).toBe(false);
      expect(ops.some((o) => o.op === 'scale')).toBe(false);
      expect(ops.some((o) => o.op === 'translate')).toBe(false);
      const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
      expect(fill).toMatchObject({ text: ch, x: 105, y: 200, align: 'center', baseline: 'middle' });
    }
  });

  it('substitutes a Tr bracket with its vertical form (（→︵, ）→︶) and draws it UPRIGHT', () => {
    // Symptom 2: a rotated fullwidth bracket lands its ink off-cell (unmeasurable
    // on a Canvas). Substituting the U+FE35/FE36 vertical form and drawing it
    // upright (counter-rotated) lets a per-glyph vertical-ink metric centre it.
    const { ctx, ops } = mockCtx({
      // Model ︵/︶ ink hugging opposite cell ends under a `middle` baseline.
      inkMiddle: { '︵': { asc: -9, desc: 21 }, '︶': { asc: 21, desc: -9 } },
    });
    drawVerticalRun(ctx, '（）', 0, 0, 12, 0);
    // Both are substituted to their vertical forms and COUNTER-ROTATED (upright).
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.text)).toEqual(['︵', '︶']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(rotates.every((r) => Math.abs(r.a - -Math.PI / 2) < 1e-9)).toBe(true);
    // Along-column ink centring: fillText y = (asc − desc)/2. For ︵: (−9−21)/2 =
    // −15; for ︶: (21−(−9))/2 = 15. Drawn centred (x = 0, the cell centre).
    expect(fills.map((f) => f.align)).toEqual(['center', 'center']);
    expect(fills.map((f) => f.baseline)).toEqual(['middle', 'middle']);
    expect(fills.map((f) => f.x)).toEqual([0, 0]);
    expect(fills[0].y).toBeCloseTo(-15, 6);
    expect(fills[1].y).toBeCloseTo(15, 6);
  });

  it('keeps a Tr bracket at the cell centre when the Canvas reports no ink metrics', () => {
    // No actualBoundingBox* → along-column correction is 0, so the substituted
    // vertical form draws at the cell centre exactly (graceful degradation).
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '「」', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.text)).toEqual(['﹁', '﹂']); // FE41 / FE42
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });

  it('gives the colon/semicolon a geometric fallback and substitutes the lenticular brackets (issue #969)', () => {
    // FE13/FE14 (vertical colon/semicolon) are absent from most render fonts and a
    // Canvas cannot invoke the font's `vert` feature, so unconditional substitution
    // reached a mispositioned system-fallback glyph. They now take a GEOMETRIC
    // fallback that reproduces each vertical form's design directly (Word-verified):
    //   • colon ： → ROTATE (drawn as-is in the +90° page frame → FE13's side-by-side
    //     dots); no local counter-rotation.
    //   • semicolon ； → UPRIGHT (counter-rotated −90° → FE14's dot-over-comma; a
    //     rotation could not produce that).
    // The white lenticular brackets 〖〗 keep their FE17/FE18 substitute (present in
    // the substitute fonts), drawn upright.
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '：；〖〗', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((f) => f.text)).toEqual(['：', '；', '︗', '︘']); // colon/semicolon base; FE17/FE18
    // Counter-rotations (−90°, upright): semicolon + both lenticular brackets = 3;
    // the colon is NOT counter-rotated (it rides the +90° page rotation).
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(3);
    expect(rotates.every((r) => Math.abs(r.a - -Math.PI / 2) < 1e-9)).toBe(true);
    expect(fills.every((f) => f.align === 'center' && f.baseline === 'middle')).toBe(true);
  });

  it('substitutes a Tu comma/period with its vertical presentation form (、→︑, 。→︒)', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '、。，', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // U+FE11 / U+FE12 / U+FE10 keep their designed position.
    expect(fills.map((f) => f.text)).toEqual(['︑', '︒', '︐']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(3);
    // Substituted forms are pre-positioned by the font → drawn at the cell centre
    // with no upper-right nudge (local x offset 0). With no ink metrics the mock
    // reports none, so the along-column correction is 0 (y = 0) too.
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });

  it('does NOT ink-centre a substituted Tu comma even when ink metrics are present', () => {
    // The comma/full stop vertical forms (︑ ︒) are DESIGNED with their ink in the
    // cell's upper-right corner (JIS X 4051). Unlike a Tr bracket, they must NOT be
    // pulled to the geometric cell centre by the along-column ink metric — doing so
    // drops them LOW (the "、。 sit too low" defect, #771). So even with ink metrics
    // reported, the punctuation draws at the em-box centre (y = 0), letting the
    // font's designed corner offset stand.
    const { ctx, ops } = mockCtx({
      // Model ︑ ink hugging the top of the cell (as a real vertical comma does).
      inkMiddle: { '︑': { asc: 21, desc: -9 } },
    });
    drawVerticalRun(ctx, '、', 0, 0, 12, 0);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill?.text).toBe('︑');
    // y stays 0 (em-box centred) — the ink metric is deliberately ignored here.
    expect(fill?.y).toBe(0);
    // Contrast: a Tr bracket WITH the same ink metric IS shifted (see the ︵/︶ test
    // above), proving the skip is punctuation-specific, not a blanket disable.
  });

  it('draws ！ / ？ UPRIGHT without substitution (they stand centred, not corner-hung)', () => {
    // ！ FF01 and ？ FF1F are vo=Tu but NOT substituted to FE15/FE16: those vertical
    // forms are corner-designed in many fonts and would push the mark off the column
    // centre (the sample-26 "！ shifted right" defect, #771). The original fullwidth
    // mark drawn upright is already the correct, centred vertical appearance.
    const { ctx, ops } = mockCtx();
    drawVerticalRun(ctx, '！？', 0, 0, 12, 0);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    // Drawn as the ORIGINAL code points (no FE15/FE16 substitution), counter-rotated
    // upright and centred on the cell (center/middle at x=0, y=0).
    expect(fills.map((f) => f.text)).toEqual(['！', '？']);
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(2);
    expect(fills.every((f) => f.align === 'center' && f.baseline === 'middle')).toBe(true);
    expect(fills.every((f) => f.x === 0 && f.y === 0)).toBe(true);
  });
});

describe('reachable vert glyph cells (issue #1024 — feature-state measure == paint)', () => {
  it('preserves contextual shaping across a consecutive sideways piece', () => {
    const text = 'Word';
    const { ctx } = mockCtx({ wholeWidths: { [text]: 35 } });

    // A vo=R piece stays sideways in the rotated frame, so its vertical advance
    // is the same contextually-shaped width as ordinary horizontal text. In
    // particular, kerning must not be replaced by four independent glyph cells.
    expect(verticalRunInkExtraPxWithCapability(ctx, text, () => false)).toBe(0);

    const painted = mockCtx({ wholeWidths: { [text]: 35 } });
    drawVerticalRunWithCapability(painted.ctx, text, 0, 0, 12, 0, 1, true, () => false);
    const fills = painted.ops.filter(
      (op): op is Extract<Op, { op: 'fillText' }> => op.op === 'fillText',
    );
    expect(fills).toHaveLength(1);
    expect(fills[0]).toMatchObject({ text, x: 0, feature: 'normal' });
  });

  it('keeps the featured origin at half-advance and allows designed leading ink to poke', () => {
    const { ctx, ops } = mockCtx({ vert: { 'ー': { asc: 8, desc: 5 } } });
    expect(verticalRunInkExtraPxWithCapability(ctx, 'ーA', () => true)).toBe(0);

    drawVerticalRunWithCapability(ctx, 'ーA', 0, 0, 12, 0, 1, true, () => true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    // Featured ink [5-8,5+5]=[-3,10] intentionally pokes before [0,10], but
    // the font's origin stays at the nominal half-advance instead of shifting.
    expect(translates[0]).toEqual({ op: 'translate', x: 5, y: 0 });
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills[0]).toMatchObject({ text: 'ー', x: 0, y: 0, feature: '"vert" 1' });
    expect(fills[1]).toMatchObject({ text: 'A', x: 10, feature: 'normal' });
  });

  it('recovers a whole-run horizontal kern deficit with the per-glyph vertical cell sum', () => {
    const text = '、。「」ー';
    const { ctx } = mockCtx({ wholeWidths: { [text]: 40 } });
    // Five independent 10px cells paint as 50px, while horizontal measureText
    // compresses the whole string to 40px. The vertical delta restores 10px.
    expect(verticalRunInkExtraPxWithCapability(ctx, text, () => true)).toBe(10);

    const painted = mockCtx({ wholeWidths: { [text]: 40 } });
    drawVerticalRunWithCapability(
      painted.ctx,
      `${text}c`,
      0,
      0,
      12,
      0,
      1,
      true,
      (cp) => cp !== 'c'.codePointAt(0),
    );
    const latin = painted.ops
      .filter((op): op is Extract<Op, { op: 'fillText' }> => op.op === 'fillText')
      .at(-1);
    expect(latin).toMatchObject({ text: 'c', x: 50, feature: 'normal' });
  });

  it('preserves complementary bracket A/D placement at one-em glyph origins', () => {
    const { ctx, ops } = mockCtx({
      vert: {
        '「': { asc: -1, desc: 4 },
        '」': { asc: 4, desc: -1 },
      },
    });
    drawVerticalRunWithCapability(ctx, '「」', 0, 0, 12, 0, 1, true, () => true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates).toEqual([
      { op: 'translate', x: 5, y: 0 },
      { op: 'translate', x: 15, y: 0 },
    ]);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => [fill.text, fill.y, fill.feature])).toEqual([
      ['「', 0, '"vert" 1'],
      ['」', 0, '"vert" 1'],
    ]);
    // The renderer does not ink-centre either glyph: the paired ink centres are
    // 0.5em apart even though their feature glyph origins remain 1em apart.
    const firstInkCenter = translates[0].x + (4 - -1) / 2;
    const secondInkCenter = translates[1].x + (-1 - 4) / 2;
    expect(secondInkCenter - firstInkCenter).toBe(5);
  });

  it('keeps the feature-designed upper-right placement of ．、。，', () => {
    const { ctx, ops } = mockCtx({
      vert: {
        '．': { asc: 8, desc: -3 },
        '、': { asc: 9, desc: -4 },
        '。': { asc: 9, desc: -4 },
        '，': { asc: 7, desc: -2 },
      },
    });
    drawVerticalRunWithCapability(ctx, '．、。，', 0, 0, 12, 0, 1, true, () => true);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => fill.text)).toEqual(['．', '、', '。', '，']);
    expect(fills.every((fill) => fill.y === 0 && fill.feature === '"vert" 1')).toBe(true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates.map((op) => op.x)).toEqual([5, 15, 25, 35]);
  });

  it('scales the featured cell and origin before adding letter spacing', () => {
    const { ctx, ops } = mockCtx({ vert: { 'ー': { asc: 8, desc: 5 } } });
    drawVerticalRunWithCapability(ctx, 'ーA', 0, 0, 12, 4, 0.5, true, () => true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates[0]).toEqual({ op: 'translate', x: 2.5, y: 0 });
    expect(translates[1]).toEqual({ op: 'translate', x: 9, y: 0 });
    // 10px featured cell * 0.5 + 4px spacing precedes the Latin cell.
    expect(ops).toContainEqual({ op: 'scale', sx: 1, sy: 0.5 });
  });

  it('does not route vo=U or vo=R through vert even when the capability says true', () => {
    const { ctx, ops } = mockCtx();
    drawVerticalRunWithCapability(ctx, '富A', 0, 0, 12, 0, 1, true, () => true);
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => fill.feature)).toEqual(['normal', 'normal']);
  });
});

// issue #1014 — a vo=Tr rotate-fallback mark (ー, 〜, quotes, colon) whose
// substitute font UNDER-REPORTS its advance via measureText draws with ink that
// spills PAST the advance-sized cell into the following sideways run (Chrome).
// The fix sizes the rotate cell to the along-column INK extent (a NO-OP unless
// the ink exceeds the advance) and ink-centres the glyph in the grown cell, so
// the mark stays inside its own cell and the next run clears it. measure==paint:
// the SAME per-glyph deficit is added to the layout advance via
// verticalRunInkExtraPx.
describe('vo=Tr rotate-fallback ink overrun (#1014 — ink-sized cell + ink-centring)', () => {
  // A `middle`/`center` measureText for ー reports advance 10 (mock: 1 cp × 10)
  // but a HORIZONTAL ink of left+right = 5+24 = 29 > 10 — the under-report.
  const underReport = { inkLR: { ー: { left: 5, right: 24 } } };

  it('verticalRunInkExtraPx sums the per-glyph ink deficit over Tr rotate glyphs only', () => {
    const { ctx } = mockCtx(underReport);
    // ー: max(0, 29 − 10) = 19. Upright 話 and sideways A/space contribute 0.
    expect(verticalRunInkExtraPx(ctx, 'ー')).toBeCloseTo(19, 6);
    expect(verticalRunInkExtraPx(ctx, '話ー')).toBeCloseTo(19, 6);
    expect(verticalRunInkExtraPx(ctx, '話A ')).toBe(0);
  });

  it('keeps quote/colon geometric growth when only the long mark is reachable', () => {
    const { ctx } = mockCtx({
      inkLR: {
        'ー': { left: 5, right: 24 },
        '“': { left: 5, right: 24 },
        '：': { left: 5, right: 24 },
      },
    });
    expect(
      verticalRunInkExtraPxWithCapability(ctx, 'ー“：', (cp) => cp === 0x30fc),
    ).toBeCloseTo(38, 6);
  });

  it('uses the same per-code-point vert gate for measurement and painting', () => {
    const metrics = {
      inkLR: {
        'ー': { left: 5, right: 24 },
        '〜': { left: 5, right: 24 },
      },
    };
    const supported = (cp: number) => cp === 0x30fc;
    const { ctx } = mockCtx(metrics);
    expect(verticalRunInkExtraPxWithCapability(ctx, 'ー〜', supported)).toBeCloseTo(19, 6);

    const painted = mockCtx(metrics);
    drawVerticalRunWithCapability(painted.ctx, 'ー〜', 0, 0, 12, 0, 1, true, supported);
    const fills = painted.ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills.map((fill) => fill.feature)).toEqual(['"vert" 1', 'normal']);
    const transforms = painted.ops.filter(
      (o): o is Extract<Op, { op: 'transform' }> => o.op === 'transform',
    );
    expect(transforms).toEqual([{ op: 'transform', a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }]);
  });

  it('applies the same per-glyph growth gate while painting', () => {
    const { ctx, ops } = mockCtx({
      inkLR: {
        'ー': { left: 5, right: 24 },
        '“': { left: 5, right: 24 },
        '：': { left: 5, right: 24 },
      },
    });
    drawVerticalRunWithCapability(
      ctx,
      'ー“：話',
      0,
      0,
      12,
      0,
      1,
      true,
      (cp) => cp === 0x30fc,
    );
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates.map((op) => op.x)).toEqual([5, 24.5, -9.5, 53.5, -9.5, 73]);
  });

  it('verticalRunInkExtraPx is 0 when the ink fits the advance (all real fonts) or metrics are absent', () => {
    const fits = mockCtx({ inkLR: { ー: { left: 3, right: 4 } } }); // extent 7 ≤ 10
    expect(verticalRunInkExtraPx(fits.ctx, 'ー')).toBe(0);
    const noMetrics = mockCtx(); // no actualBoundingBox* → graceful 0
    expect(verticalRunInkExtraPx(noMetrics.ctx, 'ー')).toBe(0);
  });

  it('grows the ー cell to its ink extent so the FOLLOWING glyph clears the ink', () => {
    const { ctx, ops } = mockCtx(underReport);
    // ー (grown cell 29) then upright 話 (advance 10). Without the fix 話 would
    // centre at 10 + 5 = 15 (inside the ー ink); with the ink-sized cell it
    // centres at 29 + 5 = 34. `growTrRotateInk=true` = the body path (whose layout
    // advance was grown by the same deficit — measure==paint).
    drawVerticalRun(ctx, 'ー話', 0, 0, 12, 0, 1, true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    // First translate = ー cell centre (29/2 = 14.5); second is its separate
    // output-axis ink shift; third = 話 cell centre 34.
    expect(translates[0].x).toBeCloseTo(14.5, 6);
    expect(translates[1].x).toBeCloseTo(-9.5, 6);
    expect(translates[2].x).toBeCloseTo(34, 6);
  });

  it('ink-centres the grown ー (shift by (left − right)/2) so its ink fills the grown cell', () => {
    const { ctx, ops } = mockCtx(underReport);
    drawVerticalRun(ctx, 'ー', 0, 0, 12, 0, 1, true);
    // Plain rotation path: translate to cell centre 14.5, then translate the
    // output advance axis by (5 − 24)/2 = −9.5 before the identity scale matrix so
    // the ink centres on the
    // cell. Ink then spans [14.5 − 9.5 − 5, 14.5 − 9.5 + 24] = [0, 29] = the cell.
    const translate = ops.find((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translate?.x).toBeCloseTo(14.5, 6);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates[1]?.x).toBeCloseTo(-9.5, 6);
    const transform = ops.find((o): o is Extract<Op, { op: 'transform' }> => o.op === 'transform');
    expect(transform).toEqual({ op: 'transform', a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill).toMatchObject({ text: 'ー', y: 0, align: 'center', baseline: 'middle' });
    expect(fill?.x).toBe(0);
  });

  it('keeps charScale and rotateInkShiftPx without reflection or shear', () => {
    const { ctx, ops } = mockCtx(underReport);
    drawVerticalRun(ctx, 'ー', 0, 0, 12, 0, 0.5, true);
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates).toEqual([
      { op: 'translate', x: 7.25, y: 0 },
      { op: 'translate', x: -4.75, y: 0 },
    ]);
    const transform = ops.find((o): o is Extract<Op, { op: 'transform' }> => o.op === 'transform');
    expect(transform).toMatchObject({ a: 0.5, b: 0, c: 0, d: 1, e: 0, f: 0 });
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill).toMatchObject({ text: 'ー', x: 0, y: 0 });
  });

  it('is a NO-OP (byte-identical) when the ink fits the advance — no growth, no shift', () => {
    // extent 7 ≤ advance 10 → the ー draws exactly as today even with grow enabled:
    // cell 10, centre 5, plain fillText in the page frame.
    const { ctx, ops } = mockCtx({ inkLR: { ー: { left: 3, right: 4 } } });
    drawVerticalRun(ctx, 'ー', 100, 200, 12, 0, 1, true);
    expect(ops.some((o) => o.op === 'translate')).toBe(false);
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill).toMatchObject({ text: 'ー', x: 105, y: 200 });
  });

  it('does NOT grow when growTrRotateInk is false (marker / unwired text box) — paint stays coupled to the measure', () => {
    // The same under-reporting ー, but growTrRotateInk defaults to false: the cell
    // stays advance-sized (10) and advance-centred (cx=5), byte-identical to the
    // pre-#1014 draw. Callers whose LAYOUT advance was NOT grown (no s.verticalRun —
    // list markers, eaVert text boxes) pass false so paint never exceeds measure.
    const { ctx, ops } = mockCtx(underReport);
    drawVerticalRun(ctx, 'ー話', 0, 0, 12, 0); // 7 args → growTrRotateInk = false
    const translates = ops.filter((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translates).toHaveLength(1);
    expect(translates[0].x).toBeCloseTo(15, 6); // only upright 話 translates
    const fill = ops.find((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fill?.x).toBe(5); // ー is plain-rotated at its advance centre
  });
});

describe('drawTateChuYokoRun (§17.3.2.10 — horizontal-in-vertical / 縦中横)', () => {
  it('draws the whole run as ONE upright, centred fillText in the cell centre', () => {
    // Two full-width digits "２９" in a 12px cell (cellAdvance=12), no scale.
    const { ctx, ops } = mockCtx();
    drawTateChuYokoRun(ctx, '２９', 100, 200, 12, 12, 1, false);
    // Exactly one fillText — the whole run, not per glyph.
    const fills = ops.filter((o): o is Extract<Op, { op: 'fillText' }> => o.op === 'fillText');
    expect(fills).toHaveLength(1);
    expect(fills[0].text).toBe('２９');
    // Centred (center/middle) at local origin.
    expect(fills[0]).toMatchObject({ align: 'center', baseline: 'middle', x: 0, y: 0 });
    // Counter-rotated −90° (upright, cancels the +90° page rotation).
    const rotates = ops.filter((o): o is Extract<Op, { op: 'rotate' }> => o.op === 'rotate');
    expect(rotates).toHaveLength(1);
    expect(rotates[0].a).toBeCloseTo(-Math.PI / 2, 12);
    // Pivot = cell centre along the column: x = 100 + cellAdvance/2 = 106, y=200.
    const translate = ops.find((o): o is Extract<Op, { op: 'translate' }> => o.op === 'translate');
    expect(translate).toEqual({ op: 'translate', x: 106, y: 200 });
  });

  it('applies w:w only to the cross-column (glyph width) axis — scale(charScale, 1)', () => {
    // §17.3.2.43 w:w=67 compresses the digits side-by-side ACROSS the column
    // (local x), NOT the along-column cell height (local y stays 1).
    const { ctx, ops } = mockCtx();
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 0.67, false);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 0.67, sy: 1 });
  });

  it('leaves the along-column axis at 1 when vertCompress content already fits one em', () => {
    // vertCompress on a single-line run whose natural height (asc+desc) equals the
    // em: no compression needed, so scale y stays 1 (the common 2-digit case).
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 3 }); // 12 = 1em
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 1, sy: 1 });
  });

  it('compresses the along-column axis when vertCompress content is taller than one em', () => {
    // A run whose natural upright height exceeds one em is squeezed to fit one
    // cell (§17.3.2.10): scale y = fontPx / height. Here height = 18 > 12 → 12/18.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 14, fontBoundingBoxDescent: 4 }); // 18px
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale?.sx).toBe(1);
    expect(scale?.sy).toBeCloseTo(12 / 18, 12);
  });

  it('does not compress the height when vertCompress is off, even if content is tall', () => {
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 14, fontBoundingBoxDescent: 4 });
    drawTateChuYokoRun(ctx, '２９', 0, 0, 12, 12, 1, false);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale).toEqual({ op: 'scale', sx: 1, sy: 1 });
  });

  it('combines w:w (width) and vertCompress (height) independently', () => {
    // 3-digit-style case: w:w=0.5 across the column AND a tall run compressed to
    // one em along the column — the two axes are set independently.
    const { ctx, ops } = mockCtx({ fontBoundingBoxAscent: 18, fontBoundingBoxDescent: 6 }); // 24px
    drawTateChuYokoRun(ctx, '２９９', 0, 0, 12, 12, 0.5, true);
    const scale = ops.find((o): o is Extract<Op, { op: 'scale' }> => o.op === 'scale');
    expect(scale?.sx).toBe(0.5);
    expect(scale?.sy).toBeCloseTo(12 / 24, 12);
  });
});

describe('drawUprightBox (§17.6.20 — keep images upright inside the rotated page)', () => {
  it('rotates −90° about the box centre and passes the swapped local box', () => {
    const { ctx, ops } = mockCtx();
    let called: number[] | null = null;
    drawUprightBox(ctx, 10, 20, 100, 40, (dx, dy, dw, dh) => {
      called = [dx, dy, dw, dh];
    });
    expect(ops).toContainEqual({ op: 'translate', x: 60, y: 40 }); // centre (10+50, 20+20)
    expect(ops).toContainEqual({ op: 'rotate', a: -Math.PI / 2 });
    // Local box: width↔height swap, centred on the pivot → (−h/2, −w/2, h, w).
    expect(called).toEqual([-20, -50, 40, 100]);
    // Balanced save/restore.
    expect(ops[0]).toEqual({ op: 'save' });
    expect(ops[ops.length - 1]).toEqual({ op: 'restore' });
  });
});
