import { describe, expect, it } from 'vitest';
import type { ParagraphBorders } from '../types.js';
import { bottomBorderExtentPt, topBorderExtentPt, wordParagraphBorderThicknessPt } from './paragraph-border-adjacency.js';

describe('paragraph border flow extent', () => {
  it('reserves spacing plus the whole visible bottom stroke (WORD_PARAGRAPH_BORDER_FLOW_RESERVATION)', () => {
    const borders = {
      top: null,
      right: null,
      bottom: { style: 'single', width: 2, space: 3 },
      left: null,
      between: null,
    } as ParagraphBorders;

    expect(bottomBorderExtentPt(borders)).toBe(5);
    expect(bottomBorderExtentPt(borders, { suppressBottom: true })).toBe(0);
    expect(bottomBorderExtentPt({
      ...borders,
      bottom: { style: 'none', width: 2, space: 3 },
    } as ParagraphBorders)).toBe(0);
    expect(bottomBorderExtentPt(null)).toBe(0);
  });

  it('reserves only a visible, selected top edge', () => {
    const borders = {
      top: { style: 'single', width: 2, space: 3 },
      between: { style: 'single', width: 1, space: 2 },
      bottom: null, left: null, right: null,
    } as ParagraphBorders;
    expect(topBorderExtentPt(borders, 'top')).toBe(5);
    // A between border takes its spacing on both sides of its stroke.
    expect(topBorderExtentPt(borders, 'between')).toBe(5);
    expect(topBorderExtentPt(borders, 'none')).toBe(0);
    expect(topBorderExtentPt({ ...borders, top: { style: 'nil', width: 2, space: 3 } } as ParagraphBorders, 'top')).toBe(0);
  });
});

describe('wordParagraphBorderThicknessPt (WORD_PARAGRAPH_BORDER_FLOW_RESERVATION)', () => {
  it.each([
    ['single', 1.5, 1.5], ['dotted', 1.5, 1.5], ['outset', 3, 3],
    ['double', 1.5, 4.5], ['triple', 3, 15],
    ['thinThickSmallGap', 0.5, 2], ['thickThinSmallGap', 6, 7.5], ['thinThickThinSmallGap', 1.5, 4.5],
    ['thinThickMediumGap', 3, 6], ['thinThickThinMediumGap', 6, 18],
    ['thickThinLargeGap', 0.5, 2.75], ['thinThickThinLargeGap', 6, 15],
    ['threeDEmboss', 0.5, 2], ['threeDEngrave', 3, 6], ['threeDEmboss', 6, 9],
    ['wave', 1.5, 3.75], ['wave', 3, 3], ['doubleWave', 0.5, 5.25],
  ] as const)('gives %s at %spt a %spt band', (style, width, thickness) => {
    expect(wordParagraphBorderThicknessPt(style, width)).toBeCloseTo(thickness, 10);
  });

  it('keeps a wave band independent of w:sz except at w:sz 12', () => {
    for (const sz of [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 24, 36, 48]) {
      expect(wordParagraphBorderThicknessPt('wave', sz / 8)).toBe(sz === 12 ? 3.75 : 3);
      expect(wordParagraphBorderThicknessPt('doubleWave', sz / 8)).toBe(5.25);
    }
  });
});
