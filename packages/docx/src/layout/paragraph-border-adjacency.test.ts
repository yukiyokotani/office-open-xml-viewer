import { describe, expect, it } from 'vitest';
import type { ParagraphBorders } from '../types.js';
import { bottomBorderExtentPt, topBorderExtentPt } from './paragraph-border-adjacency.js';

describe('paragraph border flow extent', () => {
  it('reserves spacing plus half the visible bottom stroke', () => {
    const borders = {
      top: null,
      right: null,
      bottom: { style: 'single', width: 2, space: 3 },
      left: null,
      between: null,
    } as ParagraphBorders;

    expect(bottomBorderExtentPt(borders)).toBe(4);
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
    expect(topBorderExtentPt(borders, 'top')).toBe(4);
    expect(topBorderExtentPt(borders, 'between')).toBe(2.5);
    expect(topBorderExtentPt(borders, 'none')).toBe(0);
    expect(topBorderExtentPt({ ...borders, top: { style: 'nil', width: 2, space: 3 } } as ParagraphBorders, 'top')).toBe(0);
  });
});
