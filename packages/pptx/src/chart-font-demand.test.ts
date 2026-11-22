import { afterEach, describe, expect, it, vi } from 'vitest';
import { CHART_SYNTHESIZED_TEXT } from '@silurus/ooxml-core/internal/font-scalar-demand';
import { collectSlideFontDemand } from './font-preload-demand.js';
import type { Slide } from './types';

const chartSlide = (chart: Record<string, unknown>): Slide => ({
  index: 0, slideNumber: 1, background: null,
  elements: [{ type: 'chart', x: 0, y: 0, width: 1, height: 1, rotation: 0, flipH: false, flipV: false, chart }],
} as unknown as Slide);

const scalars = (text: string) => [...text].map((ch) => ch.codePointAt(0)!);

afterEach(() => { vi.unstubAllGlobals(); });

describe('PPTX chart font demand', () => {
  it('demands authored chart strings and generated ASCII, but not face names', () => {
    const demand = collectSlideFontDemand(chartSlide({
      chartType: 'bar',
      title: '売上',
      titleFontFace: '游ゴシック',
      categories: ['東京', 'Øsaka'],
      series: [{ name: 'Δ series', values: [1, 2], valFormatCode: '#,##0"円"' }],
    }), new Set(), Number.MAX_SAFE_INTEGER);
    expect(demand).not.toBe('all');
    expect(demand).toEqual(expect.arrayContaining(scalars('売上東京ØΔ円' + CHART_SYNTHESIZED_TEXT)));
    expect(demand).not.toContain('游'.codePointAt(0));
  });

  it('declines a chart whose painted texts are not provable', () => {
    expect(collectSlideFontDemand(chartSlide({ chartType: 'regionMap', series: [] }), new Set(), Number.MAX_SAFE_INTEGER)).toBe('all');
    const builtin14 = { chartType: 'bar', dataTable: {}, series: [{ catFormatBuiltinId: 14 }] };
    vi.stubGlobal('navigator', { language: 'ja-JP-u-ca-japanese' });
    expect(collectSlideFontDemand(chartSlide(builtin14), new Set(), Number.MAX_SAFE_INTEGER)).toBe('all');
  });
});
