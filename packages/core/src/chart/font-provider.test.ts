import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import { chartFontFamilies, chartFontFamily } from './renderer.js';

describe('chart provider fonts', () => {
  it('keeps the authored face before its private alias', () => {
    const chart = {
      themeMinorFontLatin: 'Aptos',
      providerFontRoutes: { aptos: '__private_aptos' },
    } as unknown as ChartModel;

    expect(chartFontFamily(chart, null, 'minor'))
      .toBe('"Aptos", "__private_aptos", Calibri, Arial, sans-serif');
  });

  it('collects concrete theme, element, override, and rich-run faces', () => {
    const chart = {
      themeMajorFontLatin: 'Aptos Display',
      themeMinorFontLatin: 'Aptos',
      titleFontFace: '+mj-lt',
      legendFontFace: 'Legend Face',
      legendEntries: [{ fontFace: 'Override Face' }],
      titleRichRuns: [{ fontFace: 'Rich Face' }],
    } as unknown as ChartModel;

    expect(chartFontFamilies(chart).sort()).toEqual([
      'Aptos',
      'Rich Face',
      'Legend Face',
      'Aptos Display',
      'Override Face',
    ].sort());
  });
});
