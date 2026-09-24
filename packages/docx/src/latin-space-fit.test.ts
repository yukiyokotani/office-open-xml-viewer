import { describe, expect, it } from 'vitest';
import { createCanvasFontRoute } from '@silurus/ooxml-core';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';

const route = createCanvasFontRoute('Synthetic Latin', 'registered');

function context(): CanvasRenderingContext2D {
  return {
    font: '10px Synthetic Latin', letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      const width = [...text].reduce((sum, char) => sum + (char === ' ' ? 4 : 2), 0);
      return { width, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function word(text: string, ratio = 0.4, enabled = true): LayoutTextSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Synthetic Latin', fontRoute: route,
    vertAlign: null, measuredWidth: 0,
    latinSpaceAverageWidthRatio: ratio,
    latinSpaceCompressionEligible: enabled ? true : undefined,
  };
}

function lay(words: LayoutTextSeg[], width: number) {
  return layoutLines(context(), words, width, 0, 1);
}

describe('selected-face Latin inter-word space fit', () => {
  it('admits one and two-gap first-fit boundaries and conserves retained widths', () => {
    const one = lay([word('A '), word('B')], 6);
    expect(one).toHaveLength(1);
    expect((one[0]!.segments[0] as LayoutTextSeg).measuredWidth).toBe(4);
    expect(one[0]!.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0)).toBe(6);

    const two = lay([word('A '), word('B '), word('C')], 10);
    expect(two).toHaveLength(1);
    expect(two[0]!.segments.map((segment) => segment.measuredWidth)).toEqual([4, 4, 2]);
    expect(two[0]!.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0)).toBe(10);
    expect(lay([word('A '), word('B')], 5.9)).toHaveLength(2);
  });

  it('preserves natural spaces when compression is disabled or below the floor', () => {
    expect(lay([word('A ', 0.4, false), word('B', 0.4, false)], 6)).toHaveLength(2);
    expect(lay([word('A ', 1), word('B', 1)], 6)).toHaveLength(2);
  });

  it('declines a mixed-face line', () => {
    const other = { ...word('B'), fontRoute: createCanvasFontRoute('Other', 'registered') };
    expect(lay([word('A '), other], 6)).toHaveLength(2);
  });

  it('does not apply horizontal U+0020 fitting to upright vertical runs', () => {
    const vertical = [
      { ...word('A '), verticalRun: true },
      { ...word('B'), verticalRun: true },
    ];
    const lines = layoutLines(context(), vertical, 6, 0, 1, [], undefined, {}, 0,
      DEFAULT_KINSOKU_RULES, undefined, 36, 6, false, false, false,
      undefined, 'bounded', {
        fingerprint: 'vertical-zero-extra', measureRunInkExtra: () => 0,
        planRun: () => [],
      });
    expect(lines).toHaveLength(2);
  });

  it('keeps a linesAndChars grid pitch separate from the xAvg space floor', () => {
    const a = { ...word('A '), widthBalanceGridDeltaFactor: 0.5 as const };
    const b = { ...word('B'), widthBalanceGridDeltaFactor: 0.5 as const };
    const grid = { type: 'linesAndChars' as const, charSpacePt: -1, linePitchPt: 12 };
    const lines = layoutLines(context(), [a, b], 4.5, 0, 1, [], undefined, {}, 0,
      DEFAULT_KINSOKU_RULES, grid);
    expect(lines).toHaveLength(1);
    expect(lines[0]!.segments.map((segment) => segment.measuredWidth)).toEqual([3, 1.5]);
  });
});
