import { describe, it, expect } from 'vitest';
import { svgExtents, recolorSvg } from './mathjax';

// The self-contained engine conversion is covered by engine.worker.test.ts;
// these cover only the pure SVG helpers.

describe('svgExtents', () => {
  it('parses the viewBox into baseline-relative em extents', () => {
    const e = svgExtents('<svg viewBox="0 -1642.5 9178 2338.5"></svg>');
    expect(e.widthEm).toBeCloseTo(9.178, 3);
    expect(e.ascentEm).toBeCloseTo(1.6425, 3);
    expect(e.descentEm).toBeCloseTo(0.696, 3);
  });

  it('clamps an above-baseline box to zero descent', () => {
    const e = svgExtents('<svg viewBox="0 -900 1200 700"></svg>');
    expect(e).toEqual({ widthEm: 1.2, ascentEm: 0.9, descentEm: 0 });
  });

  it('clamps a below-baseline box to zero ascent', () => {
    const e = svgExtents('<svg viewBox="0 200 1200 700"></svg>');
    expect(e).toEqual({ widthEm: 1.2, ascentEm: 0, descentEm: 0.9 });
  });

  it('returns zeros when no viewBox is present', () => {
    expect(svgExtents('<svg></svg>')).toEqual({ widthEm: 0, ascentEm: 0, descentEm: 0 });
  });
});

describe('recolorSvg', () => {
  it('replaces currentColor', () => {
    expect(recolorSvg('fill="currentColor" stroke="currentColor"', '#f00')).toBe(
      'fill="#f00" stroke="#f00"',
    );
  });
});
