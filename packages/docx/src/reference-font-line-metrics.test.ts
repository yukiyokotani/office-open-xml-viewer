import { describe, expect, it, vi } from 'vitest';
import {
  referenceFontAverageWidthRatio,
  referenceFontLineMetrics,
} from './reference-font-line-metrics.js';

describe('DOCX reference font vertical metrics', () => {
  it('projects Calibri regular hhea geometry for automatic line spacing', () => {
    const metric = referenceFontLineMetrics('Calibri', 400, 'normal', 'other');
    expect(metric).toEqual({
      lineHeightRatio: 2500 / 2048,
      designAscentRatio: 1950 / 2048,
      designDescentRatio: 550 / 2048,
    });
    expect(11 * metric!.lineHeightRatio * (259 / 240))
      .toBeCloseTo(14.490763346354166, 12);
  });

  it('prefers macOS metadata on macOS and Office metadata elsewhere', () => {
    expect(referenceFontLineMetrics('Times New Roman', 400, 'normal', 'macos')
      ?.lineHeightRatio).toBe(2355 / 2048);
    expect(referenceFontLineMetrics('Times New Roman', 400, 'normal', 'other')
      ?.lineHeightRatio).toBe(2268 / 2048);
  });

  it('includes the macOS system Symbol face before the Office SymbolMT face', () => {
    expect(referenceFontLineMetrics('Symbol', 400, 'normal', 'macos')?.lineHeightRatio)
      .toBe((1436 + 612) / 2048);
    expect(referenceFontLineMetrics('Symbol', 400, 'normal', 'other')?.lineHeightRatio)
      .toBe((2059 + 450) / 2048);
    expect(referenceFontAverageWidthRatio('Symbol', 400, 'normal', 'macos'))
      .toBe(1172 / 2048);
    expect(referenceFontAverageWidthRatio('Symbol', 400, 'normal', 'other'))
      .toBe(1229 / 2048);
  });

  it('uses the host platform when Node has no global Navigator', async () => {
    vi.stubGlobal('navigator', undefined);
    vi.resetModules();
    try {
      const { referenceFontLineMetrics: lookup } = await import('./reference-font-line-metrics.js');
      const expected = process.platform === 'darwin' ? 2355 / 2048 : 2268 / 2048;
      expect(lookup('Times New Roman')?.lineHeightRatio).toBe(expected);
    } finally {
      vi.unstubAllGlobals();
      vi.resetModules();
    }
  });

  it('uses the Office Far-East allocation for Meiryo even on a Latin line', () => {
    const metric = referenceFontLineMetrics('Meiryo', 400, 'normal', 'other');
    const box = 2171 + 901;
    expect(metric?.lineHeightRatio).toBeCloseTo(1.3 * box / 2048, 12);
    expect(metric?.designAscentRatio).toBeCloseTo((2171 + 0.15 * box) / 2048, 12);
    expect(metric?.designDescentRatio).toBeCloseTo((901 + 0.15 * box) / 2048, 12);
  });

  it('keeps a 12pt Yu Mincho line below an 18pt grid step', () => {
    const metric = referenceFontLineMetrics('Yu Mincho', 400, 'normal', 'other');
    expect(metric?.lineHeightRatio).toBeCloseTo(1.3 * (1802 + 455) / 2048, 12);
    expect(12 * metric!.lineHeightRatio).toBeLessThan(18);
  });

  it('retains the signed ordinary hhea leading of Times New Roman', () => {
    const office = referenceFontLineMetrics('Times New Roman', 400, 'normal', 'other');
    const macos = referenceFontLineMetrics('Times New Roman', 400, 'normal', 'macos');
    expect(office?.lineHeightRatio).toBe((1825 + 443) / 2048);
    expect(macos?.lineHeightRatio).toBe((1825 + 443 + 87) / 2048);
    expect(macos?.designAscentRatio).toBe((1825 + 87) / 2048);
    const negativeGap = referenceFontLineMetrics('Tamil Sangam MN', 400, 'normal', 'macos');
    expect(negativeGap?.lineHeightRatio).toBe((1550 + 717 - 210) / 2048);
    expect(negativeGap?.designAscentRatio).toBe((1550 - 210) / 2048);
  });

  it('declines an unknown family and an unavailable style tuple', () => {
    expect(referenceFontLineMetrics('Not In Reference Catalog')).toBeUndefined();
    expect(referenceFontLineMetrics('Calibri', 900, 'italic')).toBeUndefined();
    // OS/2 code-page data is absent; a listed family alone cannot choose the
    // ordinary or Far-East Word allocation class.
    expect(referenceFontLineMetrics('AppleGothic', 400, 'normal', 'macos')).toBeUndefined();
  });
});
