import { describe, expect, it } from 'vitest';
import { distributeLineSlack } from '@silurus/ooxml-core';
import { distributedDelta } from '../text-distribute.js';
import { lineGapModel, textCellGapOptions } from './line-gaps.js';
import { wordJustifiedInterwordCompressionFactor } from '../layout/line-compatibility.js';
import { layoutLines, type LayoutTextSeg } from '../line-layout.js';

// Synthetic Word control classes: space runs, NBSP, tab/object/ruby cells,
// mixed sizes, line edges. Fit and retained paint must share opportunities.
describe('structural line gaps', () => {
  const segments = [
    { text: ' a  b\u00a0c ', widthPx: 22, spacePx: 2 },
    { widthPx: 11, spacePx: 0 },
    { text: ' d ', widthPx: 12, spacePx: 3 },
    { text: '\t', widthPx: 9, spacePx: 0 },
    { text: ' e ', widthPx: 12, spacePx: 3 },
    { text: 'f ', widthPx: 10, spacePx: 4 },
    { text: 'g ', widthPx: 10, spacePx: 2 },
  ];
  it('uses the same structural opportunities in fit and positive/negative paint', () => {
    const model = lineGapModel(segments);
    expect(model.gaps.map(gap => [gap.segIndex, gap.cpOffset, gap.naturalPx]))
      .toEqual([[0, 2, 2], [0, 3, 2], [4, 2, 3], [5, 1, 4]]);
    expect(model.S).toBe(11);
    expect(model.lineEndSeparatorPx).toBe(2);
    expect(model.visibleWidthPx).toBe(84);
    for (const slack of [-0.25, 8]) {
      const dist = distributeLineSlack(segments, slack, {
        ...textCellGapOptions, gapWidth: si => segments[si].spacePx,
      });
      expect(distributedDelta(dist)).toBeCloseTo(slack, 10);
      expect(dist!.perSeg.get(0)!.gapDeltas).toEqual([slack * 2 / 11, slack * 2 / 11]);
      expect(dist!.perSeg.get(4)!.trailingDelta).toBeCloseTo(slack * 3 / 11, 10);
      expect(dist!.perSeg.get(5)!.trailingDelta).toBeCloseTo(slack * 4 / 11, 10);
    }
  });

  it('preserves a pending space run across source seams and opaque neighbours', () => {
    let prefix = lineGapModel([], undefined, false);
    for (const segment of segments) prefix = lineGapModel([segment], prefix, false);
    const whole = lineGapModel(segments);
    expect(prefix.S).toBe(whole.S);
    expect(prefix.visibleWidthPx).toBe(whole.visibleWidthPx);
    const before = lineGapModel([{ text: 'a ', widthPx: 6, spacePx: 2 }], undefined, false);
    const after = lineGapModel([{ text: ' ', widthPx: 3, spacePx: 3 }, { text: 'b', widthPx: 4, spacePx: 2 }], before);
    expect(after.S).toBe(5);
  });

  it('handles a long space run without argument fan-out or prefix rescanning', () => {
    const spaces = ' '.repeat(50_000);
    const model = lineGapModel([{ text: `a${spaces}b`, widthPx: 100_002, spacePx: 2 }]);
    expect(model.S).toBe(100_000);
    expect(model.gaps).toHaveLength(50_000);
  });

  it('keeps a combining-space cluster fixed inside a run and across seams', () => {
    for (const segments of [
      [{ text: 'a  \u0301b c', widthPx: 24, spacePx: 2 }],
      [{ text: 'a  ', widthPx: 8, spacePx: 2 }, { text: '\u0301b c', widthPx: 16, spacePx: 2 }],
    ]) {
      const model = lineGapModel(segments);
      expect(model.S).toBe(4); // first standalone space, then the b/c gap
      let prefix = lineGapModel([], undefined, false);
      for (const segment of segments) prefix = lineGapModel([segment], prefix, false);
      expect(prefix.S).toBe(4);
      expect(model.gaps).toHaveLength(2);
    }
  });
});

it('admits a protected successor group as a whole even after a separator', () => {
  const seg = (text: string, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg => ({
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'serif', vertAlign: null, measuredWidth: 0, ...extra,
  });
  const ctx = { font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
    measureText: (text: string) => ({ width: [...text].length * 10,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
  const lines = layoutLines(ctx, [
    seg('lead '), seg('ab '),
    seg('-', { joinPrev: true, hardJoinPrev: true, noBreakRanges: [{ start: 0, end: 1 }] }),
    seg('xyz', { joinPrev: true, hardJoinPrev: true }),
  ], 70, 0, 1, [], undefined, {}, 0, undefined, undefined, undefined, undefined,
  false, true, false, undefined, 'bounded', undefined, false, true);
  expect(lines.map(line => line.segments.map(s => 'text' in s ? s.text : '').join('')))
    .toEqual(['lead ', 'ab -xyz']);
});

it('uses both observed compression boundaries, including one and zero gaps', () => {
  const input = { overflow: 0.25, naturalGapSum: 2, candidateLineEndSeparator: 2,
    previousOpportunitySum: 2, expansionWithoutCandidate: 4 };
  expect(wordJustifiedInterwordCompressionFactor(input)).toBe(0.125);
  expect(wordJustifiedInterwordCompressionFactor({ ...input, overflow: 0.51 })).toBeUndefined();
  expect(wordJustifiedInterwordCompressionFactor({ ...input, naturalGapSum: 0 })).toBeUndefined();
  // Word H7's unequal spaces bracket the second boundary independently of
  // the quarter bound. Widths/natural advances are from the control manifest.
  for (const [width, accepted] of [[441.65, false], [442.15, true]] as const) {
    const result = wordJustifiedInterwordCompressionFactor({
      overflow: 444.170898438 - width,
      naturalGapSum: 48.2421875,
      candidateLineEndSeparator: 2.653320313,
      previousOpportunitySum: 48.2421875,
      expansionWithoutCandidate: width - (444.170898438 - 3.22265625 - 2.653320313),
    });
    expect(result !== undefined).toBe(accepted);
  }
});


describe('retained contextual space projection through line acquisition', () => {
  const cluster = (start: number, advancePt: number): NonNullable<LayoutTextSeg['shapedSpaceClusters']>[number] => ({
    range: { start, end: start + 1 }, offsetPt: 0, advancePt,
  });
  function acquire(width: number, firstClusters = [cluster(3, 100), cluster(4, 3), cluster(3, 1)]) {
    // Native contextual widths: emoji8 + a6, adjacent spaces1/3;
    // b8 + its contextual space4; cccc18. Retained offsets are UTF-16,
    // whereas gap enumeration sees the astral scalar as one code point.
    const context = { font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText(text: string) {
        const widths: Record<string, number> = { '😀': 8, a: 6, b: 8, c: 4.5, ' ': 1 };
        const width = [...text].reduce((n, c) => n + (widths[c] ?? 5), 0)
          + (text.match(/  /g)?.length ?? 0) * 2 + (text.match(/b /g)?.length ?? 0) * 3;
        return { width, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const segment = (text: string, shapedSpaceClusters?: LayoutTextSeg['shapedSpaceClusters']): LayoutTextSeg => ({
      text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'serif', vertAlign: null, measuredWidth: 0,
      ...(shapedSpaceClusters ? { shapedSpaceClusters } : {}),
    });
    return layoutLines(context, [segment('😀a  ', firstClusters), segment('b ', [cluster(1, 4)]), segment('cccc')],
      width, 0, 1, [], undefined, {}, 0, undefined, undefined, undefined, undefined,
      false, true, false, undefined, 'bounded', undefined, false, { admissionOnly: true, lineEndIdeographicSpaceHangs: false });
  }
  it('keeps unequal contextual advances and the last duplicate offset across astral text', () => {
    // Natural48 and S=1+3+4=8: the existing quarter-space cap admits2pt,
    // refuses2.25pt. An isolated-space scalar, wrong offset or first duplicate
    // changes this observable boundary; no internal Map shape is asserted.
    const text = (width: number) => acquire(width).map(line => line.segments
      .map(segment => 'text' in segment ? segment.text : '').join(''));
    expect(text(46)).toEqual(['😀a  b cccc']);
    expect(text(45.75)).toEqual(['😀a  b ', 'cccc']);
  });
  it('preserves the exact quarter-cap boundary and its adjacent lower band', () => {
    const context = { font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText(text: string) {
        const widths: Record<string, number> = { a: 6, b: 8, c: 4.5, ' ': 1 };
        return { width: [...text].reduce((sum, character) => sum + widths[character]!, 0)
          + (text.includes('b ') ? 3 : 0),
          fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const segment = (text: string, spaceWidth?: number): LayoutTextSeg => ({
      text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'serif', vertAlign: null, measuredWidth: 0,
      ...(spaceWidth === undefined ? {} : { shapedSpaceClusters: [cluster(1, spaceWidth)] }),
    });
    const acquire = (width: number) => layoutLines(context,
      [segment('a ', 1), segment('b ', 4), segment('cccc')], width, 0, 1,
      [], undefined, {}, 0, undefined, undefined, undefined, undefined,
      false, true, false, undefined, 'bounded', undefined, false,
      { admissionOnly: true, lineEndIdeographicSpaceHangs: false })
      .map(line => line.segments.map(part => 'text' in part ? part.text : '').join(''));
    // Natural37pt, authoritative gaps1+4=5pt: the existing cap permits
    // exactly1.25pt. The adjacent lower double must refuse, without an epsilon.
    const bits = new BigUint64Array(new Float64Array([35.75]).buffer);
    bits[0] = bits[0]! - 1n;
    const adjacentLowerBand = new Float64Array(bits.buffer)[0]!;
    expect(acquire(35.75)).toEqual(['a b cccc']);
    expect(acquire(adjacentLowerBand)).toEqual(['a b ', 'cccc']);
  });
  it('fails closed when retained space geometry is present but empty', () => {
    expect(() => acquire(46, [])).toThrow('A text space lacks authoritative cluster geometry');
  });
});
