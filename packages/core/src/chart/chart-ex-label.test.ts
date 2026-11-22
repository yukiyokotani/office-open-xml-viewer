import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartSeries } from '../types/chart.js';
import { resolveChartExLabel } from './chart-ex-label.js';

// A label's value and percentage have separate formats in an OpenDocument
// chart (ODF 1.3 Part 3 §19.511), and the value's format may divide it
// (§19.346 `number:display-factor`).
describe('resolveChartExLabel formats', () => {
  const chart = {} as ChartModel;
  const defaults = { visible: true, showVal: true, showCatName: false };
  const label = (labels: Partial<NonNullable<ChartSeries['seriesDataLabels']>>, value: number, ratio?: number) =>
    resolveChartExLabel(
      chart,
      { name: 'S', seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false, ...labels } } as ChartSeries,
      0, 'A', value, defaults, new Map(), ratio,
    )?.text;

  it('divides the value by its format divisor', () => {
    expect(label({ formatCode: '0.0', formatDivisors: [1000] }, 2500)).toBe('2.5');
    expect(label({ formatCode: '0.0' }, 2500)).toBe('2500.0');
  });

  it('formats the percentage by its own format, else by the value format', () => {
    const both = { showPercent: true, formatCode: '0.0' };
    expect(label({ ...both, percentFormatCode: '0.0%' }, 1, 0.25)).toBe('1.0 25.0%');
    expect(label({ ...both, percentFormatCode: '0.00%', percentFormatDivisors: [1000] }, 1, 0.25)).toBe('1.0 0.03%');
    expect(label(both, 1, 0.25)).toBe('1.0 0.3');
    expect(label({ showPercent: true }, 1, 0.25)).toBe('1 25%');
  });
});
