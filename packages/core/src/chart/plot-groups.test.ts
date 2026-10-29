import { expect, test } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import { classicPlotDispatch } from './plot-groups.js';

function combo(secondaryCatHidden: boolean): ChartModel {
  return {
    chartType: 'clusteredBar',
    series: [{}, {}, {}, {}],
    secondaryCatAxis: { hidden: secondaryCatHidden },
    secondaryValAxis: { hidden: false },
    plotGroups: [
      { kind: 'bar', seriesStart: 0, seriesCount: 2, barDirection: 'col',
        categoryAxis: 'secondary', valueAxis: 'secondary', seriesAxis: 'none' },
      { kind: 'line', seriesStart: 2, seriesCount: 2,
        categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none' },
    ],
  } as unknown as ChartModel;
}

test('columns on a hidden secondary category axis combine with a primary line', () => {
  // Office draws the columns against the same categories as the line: the secondary
  // category axis is deleted, so it has no layout of its own to measure.
  expect(classicPlotDispatch(combo(true))).toBe('bar-combo');
});

test('a visible secondary category axis stays outside the combined painter', () => {
  expect(classicPlotDispatch(combo(false))).toBe('unsupported');
});
