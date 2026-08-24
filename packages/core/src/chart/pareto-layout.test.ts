import { describe, expect, it } from 'vitest';
import type { ChartSeries } from '../types/chart';
import { planParetoLayout } from './pareto-layout.js';

function series(values: Array<number | null>): ChartSeries {
  return { name: 'Frequency', color: null, values };
}

describe('planParetoLayout', () => {
  it('can retain authored order for a standalone paretoLine', () => {
    const layout = planParetoLayout(
      series([4, 8, 6, 4]),
      ['A', 'B', 'C', 'D'],
      { sortDescending: false },
    );

    expect(layout.points.map(point => point.sourceIndex)).toEqual([0, 1, 2, 3]);
    expect(layout.series.values).toEqual([4 / 22, 12 / 22, 18 / 22, 1]);
  });

  it('sorts descending, preserves tie source order, and omits invalid values', () => {
    const layout = planParetoLayout(
      series([5, 20, 10, 20, null, 0, -2, 10]),
      ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
    );

    expect(layout.points.map(point => point.sourceIndex)).toEqual([1, 3, 2, 7, 0, 5]);
    expect(layout.categories).toEqual(['B', 'D', 'C', 'H', 'A', 'F']);
    expect(layout.points.at(-1)?.cumulativeFraction).toBe(1);
  });

  it('retains all-zero values in source order without NaN or Infinity', () => {
    const layout = planParetoLayout(series([0, 0, 0]), ['A', 'B', 'C']);
    expect(layout.points.map(point => point.sourceIndex)).toEqual([0, 1, 2]);
    expect(layout.series.values).toEqual([0, 0, 0]);
    expect(layout.series.values.every(Number.isFinite)).toBe(true);
  });

  it('reaches exactly one for decimal data without rounding the running sum', () => {
    const layout = planParetoLayout(series([0.1, 1000.25, 0.2]), ['A', 'B', 'C']);
    expect(layout.series.values.at(-1)).toBe(1);
    expect(layout.series.values[0]).toBeGreaterThan(0.99);
  });

  it('keeps indexed point and label formatting attached to source identity', () => {
    const source = series([5, 20, 10]);
    source.dataPointColors = ['AAAAAA', 'BBBBBB', 'CCCCCC'];
    source.dataLabelColors = ['111111', '222222', '333333'];
    source.dataPointOverrides = [{ idx: 1, color: 'ABCDEF' }];
    source.dataLabelOverrides = [{ idx: 0, text: 'five' }];

    const layout = planParetoLayout(source, ['A', 'B', 'C']);
    expect(layout.series.dataPointColors).toEqual(['BBBBBB', 'CCCCCC', 'AAAAAA']);
    expect(layout.series.dataLabelColors).toEqual(['222222', '333333', '111111']);
    expect(layout.series.dataPointOverrides).toEqual([{ idx: 0, color: 'ABCDEF' }]);
    expect(layout.series.dataLabelOverrides).toEqual([{ idx: 2, text: 'five' }]);
    expect(layout.orderedSeries.values).toEqual([20, 10, 5]);
    expect(layout.orderedSeries.dataPointOverrides).toEqual([{ idx: 0, color: 'ABCDEF' }]);
    expect(layout.orderedSeries.dataLabelOverrides).toEqual([{ idx: 2, text: 'five' }]);
  });

  it('omits non-finite values defensively', () => {
    const layout = planParetoLayout(series([1, Number.NaN, Number.POSITIVE_INFINITY, 2]), []);
    expect(layout.points.map(point => point.sourceIndex)).toEqual([3, 0]);
    expect(layout.series.values).toEqual([2 / 3, 1]);
  });

  it('uses series-local categories when chart-level categories are absent', () => {
    const source = series([1, 3, 2]);
    source.categories = ['One', 'Three', 'Two'];
    const layout = planParetoLayout(source, []);
    expect(layout.categories).toEqual(['Three', 'Two', 'One']);
    expect(layout.series.categories).toEqual(['Three', 'Two', 'One']);
  });

  it('keeps owner-series categories authoritative when chart categories differ', () => {
    const source = series([1, 3, 2]);
    source.categories = ['Owner one', 'Owner three', 'Owner two'];
    const layout = planParetoLayout(source, ['Chart one', 'Chart three', 'Chart two']);
    expect(layout.categories).toEqual(['Owner three', 'Owner two', 'Owner one']);
  });

  it('fills sparse parallel arrays with null rather than contract-invalid undefined', () => {
    const source = series([3, 2, 1]);
    source.dataPointColors = ['AAAAAA'];
    source.dataLabelColors = [];
    source.catFormatCodes = ['0'];
    const layout = planParetoLayout(source, []);
    expect(layout.series.dataPointColors).toEqual(['AAAAAA', null, null]);
    expect(layout.series.dataLabelColors).toEqual([null, null, null]);
    expect(layout.series.catFormatCodes).toEqual(['0', null, null]);
  });

  it('keeps proportional fractions finite when finite inputs would overflow a raw sum', () => {
    const max = Number.MAX_VALUE;
    const layout = planParetoLayout(series([max, max]), ['A', 'B']);
    expect(layout.series.values).toEqual([0.5, 1]);
  });
});
