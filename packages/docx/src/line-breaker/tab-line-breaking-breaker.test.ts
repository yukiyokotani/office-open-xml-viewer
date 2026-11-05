import { describe, expect, it } from 'vitest';
import { layoutLines, type DocGridCtx, type LayoutLine, type LayoutSeg, type LayoutTabSeg, type LayoutTextSeg } from '../line-layout.js';
import type { LineFitWidth } from './pass-driver.js';
import { wordTabGlueUnit } from './tabs.js';

// The tab rules inside the line breaker, on 5 pt glyphs.

function context(): CanvasRenderingContext2D {
  return { measureText: (text: string) => ({ width: [...text].length * 5,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
}
function token(text: string, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg {
  return { text, fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0, ...extra };
}
const tab = (): LayoutSeg =>
  ({ isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false }) as unknown as LayoutSeg;

function run(segs: LayoutSeg[], options: {
  maxWidth?: number; baseRtl?: boolean; characterGrid?: DocGridCtx; lineFit?: LineFitWidth;
  tabStops?: { pos: number; alignment: string; leader: 'none' }[];
  widthPolicy?: 'bounded' | 'intrinsic';
}): LayoutLine[] {
  const maxWidth = options.maxWidth ?? 468;
  return layoutLines(context(), segs, maxWidth, 0, 1, (options.tabStops ?? []) as never, undefined, {}, 0,
    undefined, options.characterGrid, undefined, maxWidth, options.baseRtl ?? false,
    true, false, undefined, options.widthPolicy ?? 'bounded', undefined, true, false, options.lineFit);
}
const texts = (line: LayoutLine) => line.segments.flatMap((segment) => 'text' in segment ? [segment.text] : []).join('');

describe('aligned tab fit-back under the WORD_KASHIDA_LINE_FIT reserve', () => {
  it.each([false, true])('admits the gapless cell against the band, rtl=%s', (baseRtl) => {
    // highKashida (fit 0.8 of 468 pt): a 10 pt prefix, a centre stop at 400 pt
    // and a 400 pt cell, which fits the band once its gap is removed.
    const lines = run([token('ab'), tab(), token('B'.repeat(80))], {
      baseRtl, tabStops: [{ pos: 400, alignment: 'center', leader: 'none' }],
      lineFit: { fraction: 0.8, reserveShare: 0, rtlVisiblePrefix: false },
    });
    expect(lines).toHaveLength(1);
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK on a wrapped cell', () => {
  it.each([false, true])('judges the complete cell, rtl=%s', (baseRtl) => {
    // A 100 pt band, 'abcd', a centre stop at 80 pt and a 90 pt cell that
    // wraps: its first-line part alone would fit the stop, the cell does not.
    // The gap is deferred (0 at placement, the stop-aligned 15 pt at most
    // from the slack), so a justified line keeps it empty in both directions.
    const rtl = { rtl: baseRtl || undefined };
    const lines = run([token('abcd'), tab(), token('aa ', rtl), token('aa ', rtl), token('b'.repeat(12), rtl)], {
      maxWidth: 100, baseRtl, tabStops: [{ pos: 80, alignment: 'center', leader: 'none' }],
    });
    expect(lines.length).toBeGreaterThan(1);
    const placed = lines[0]!.segments.find((segment) => 'isTab' in segment) as LayoutTabSeg;
    expect(placed.measuredWidth).toBe(0);
    expect(placed.alignedGapLimit).toBe(15);
  });
});

describe('WORD_ALIGNED_TAB_CELL_FIT_BACK on a wrapped decimal cell', () => {
  it('keeps the complete LTR cell point independent of its fitted prefix', () => {
    // Stop 80 after a 20pt prefix: the complete 105pt cell's decimal point is
    // 90pt into it, on a later line. Its target precedes the prefix, so no
    // gap is allocated. RTL late-point coverage uses observable production
    // placement in rtl-tab-stops.test.ts, not a logical-prefix metadata value.
    const lines = run([token('abcd'), tab(), token('aa '), token('aa '), token('bbbbbbbbbbbb.xx')], {
      maxWidth: 100, tabStops: [{ pos: 80, alignment: 'decimal', leader: 'none' }],
    });
    expect(lines.length).toBeGreaterThan(1);
    const placed = lines[0]!.segments.find((segment) => 'isTab' in segment) as LayoutTabSeg;
    expect(placed.measuredWidth).toBe(0);
  });
});

describe('WORD_TAB_GLUED_TO_FOLLOWING_TEXT scope', () => {
  it('leaves a snapToChars word split by a formatting seam on the tab line', () => {
    // A 10 pt pitch: the ten-character word in two joined runs takes 50 pt and
    // fits the 55 pt behind a 413 pt stop.
    const grid: DocGridCtx = { type: 'snapToChars', linePitchPt: null, characterPitchPt: 10, charSpacePt: 0 };
    const lines = run([token('ab'), tab(), token('xxxxx'), token('yyyyy', { joinPrev: true })], {
      characterGrid: grid, tabStops: [{ pos: 413, alignment: 'left', leader: 'none' }],
    });
    expect(lines.map(texts)).toEqual(['abxxxxxyyyyy']);
  });

  it('adds no lookahead work for CJK text after a tab', () => {
    // Opening brackets offer no break: the rule leaves CJK out of scope, and
    // intrinsic acquisition, which never breaks automatically, skips it. A
    // per-character scan of the run made this quadratic (about 0.5 s here).
    const brackets = '\uff08'.repeat(8000);
    const time = (segs: LayoutSeg[]) => {
      const started = performance.now();
      run(segs, { widthPolicy: 'intrinsic', tabStops: [{ pos: 400, alignment: 'left', leader: 'none' }] });
      return performance.now() - started;
    };
    const alone = time([token('ab'), token(brackets)]);
    const tabbed = time([token('ab'), tab(), token(brackets)]);
    expect(wordTabGlueUnit(token(brackets), undefined, false, false)).toBe(false);
    expect(tabbed).toBeLessThan(alone * 10 + 60);
  });
});
