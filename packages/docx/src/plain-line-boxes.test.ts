import { describe, expect, it } from 'vitest';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';

describe('producer plain line boxes', () => {
  it('projects a Far East face without the Word 1.3 x usWin box when asked', () => {
    const word = referenceFontLineMetrics('MS Mincho');
    const plain = referenceFontLineMetrics('MS Mincho', 400, 'normal', true);
    expect(word?.lineHeightRatio).toBeCloseTo(1.3, 5);
    expect(plain?.lineHeightRatio).toBeCloseTo(1, 5);
  });

  it('leaves faces outside the Far East class unchanged', () => {
    expect(referenceFontLineMetrics('Times New Roman', 400, 'normal', true))
      .toEqual(referenceFontLineMetrics('Times New Roman'));
  });
});
