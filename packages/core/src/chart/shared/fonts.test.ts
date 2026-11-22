import { describe, expect, it } from 'vitest';
import { chartFaceList, chartFallbackFontStack, chartFontSubstitute } from './fonts.js';

describe('chart font substitutions', () => {
  it('inserts a loaded substitute after the authored face and the Calibri head', () => {
    const chart = { fontSubstitutions: { calibri: 'Carlito', 'brand face': 'Brand Served' } };
    expect(chartFaceList('Brand Face', chart)).toBe('"Brand Face", "Brand Served"');
    expect(chartFallbackFontStack(chart)).toBe('Calibri, "Carlito", Arial, sans-serif');
    // Every named fallback family takes the family loaded for it.
    expect(chartFallbackFontStack({ fontSubstitutions: { arial: 'ooxml-font-source-1-0' } }))
      .toBe('Calibri, Arial, "ooxml-font-source-1-0", sans-serif');
  });

  it('never resolves a document face through the object prototype', () => {
    const chart = { fontSubstitutions: { calibri: 'Carlito' } };
    for (const face of ['constructor', 'toString', 'hasOwnProperty', '__proto__', 'valueOf']) {
      expect(chartFontSubstitute(chart, face)).toBeUndefined();
      expect(chartFaceList(face, chart)).toBe(`"${face}"`);
    }
  });

  it('ignores non-string values', () => {
    const chart = { fontSubstitutions: { odd: 42 } as unknown as Record<string, string> };
    expect(chartFontSubstitute(chart, 'Odd')).toBeUndefined();
  });
});
