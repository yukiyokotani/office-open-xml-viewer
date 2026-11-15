import { describe, expect, it } from 'vitest';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';

describe('producer plain line boxes', () => {
  it('projects a Far East face without the Word 1.3 x usWin box when asked', () => {
    const word = referenceFontLineMetrics('MS Mincho');
    const plain = referenceFontLineMetrics('MS Mincho', 400, 'normal', 'plain');
    expect(word?.lineHeightRatio).toBeCloseTo(1.3, 5);
    expect(plain?.lineHeightRatio).toBeCloseTo(1, 5);
  });

  it('leaves faces outside the Far East class unchanged', () => {
    expect(referenceFontLineMetrics('Times New Roman', 400, 'normal', 'plain'))
      .toEqual(referenceFontLineMetrics('Times New Roman'));
  });
});

describe('LIBREOFFICE_FONT_LINE_BOX catalog projection', () => {
  // LibreOffice 26.8 odt-font-line-box controls, 12 pt text below a 20 pt
  // top margin: baseline and line pitch in points.
  it.each([
    // Hiragino Sans W4, hhea 880/-120/500: baseline 30.55, pitch 18.00 (Word's
    // Far East box would be 1.634 em).
    ['Hiragino Sans', 10.55, 18.0],
    // Meiryo, hhea 2171/-901/0 (Word 1.95 em): baseline 32.70, pitch 18.00.
    ['Meiryo', 12.7, 18.0],
    // MS Mincho, hhea 220/-36/0 (Word 1.3 em): baseline 30.30, pitch 12.00.
    ['MS Mincho', 10.3, 12.0],
    // Arial, hhea 1854/-434/67: baseline 30.85, pitch 13.80, the gap below.
    ['Arial', 10.85, 13.8],
  ])('puts the whole line gap of %s below the baseline', (family, ascentPt, pitchPt) => {
    const box = referenceFontLineMetrics(family, 400, 'normal', 'libreoffice');
    expect(box!.designAscentRatio * 12).toBeCloseTo(ascentPt, 1);
    expect(box!.lineHeightRatio * 12).toBeCloseTo(pitchPt, 1);
    expect(box!.designAscentRatio + box!.designDescentRatio).toBeCloseTo(box!.lineHeightRatio, 9);
  });

  it('needs no OS/2 code-page class, which only Word\'s projection reads', () => {
    expect(referenceFontLineMetrics('Courier')).toBeUndefined();
    expect(referenceFontLineMetrics('Courier', 400, 'normal', 'libreoffice')?.lineHeightRatio).toBeGreaterThan(0);
  });
});
