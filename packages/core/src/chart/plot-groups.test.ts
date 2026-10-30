import { expect, test } from 'vitest';
import type { ChartModel, ChartPlotGroup } from '../types/chart.js';
import { classicPlotDispatch } from './plot-groups.js';

type Kind = 'bar' | 'line' | 'area';

// Excel's secondary-axis combo: the first group on the primary axes, the
// second on the secondary value axis and its deleted category axis.
function combo(
  kinds: [Kind, Kind],
  secondaryCatAxis: { hidden: boolean; orientation?: string; dateAxis?: boolean },
): ChartModel {
  const group = (kind: Kind, start: number, axis: 'primary' | 'secondary'): ChartPlotGroup => ({
    kind, seriesStart: start, seriesCount: 1, grouping: kind === 'bar' ? 'clustered' : 'standard',
    ...(kind === 'bar' ? { barDirection: 'col' } : {}),
    categoryAxis: axis, valueAxis: axis, seriesAxis: 'none',
  } as ChartPlotGroup);
  return {
    chartType: kinds[0] === 'bar' ? 'clusteredBar' : kinds[0],
    series: [{}, {}],
    secondaryCatAxis,
    secondaryValAxis: { hidden: false },
    plotGroups: [group(kinds[0], 0, 'primary'), group(kinds[1], 1, 'secondary')],
  } as unknown as ChartModel;
}

test('groups on a deleted secondary category axis share the primary categories', () => {
  const hidden = { hidden: true };
  expect(classicPlotDispatch(combo(['bar', 'line'], hidden))).toBe('bar-combo');
  expect(classicPlotDispatch(combo(['line', 'bar'], hidden))).toBe('bar-combo');
  expect(classicPlotDispatch(combo(['bar', 'area'], hidden))).toBe('bar-combo');
  expect(classicPlotDispatch(combo(['line', 'line'], hidden))).toBe('line-groups');
  expect(classicPlotDispatch(combo(['area', 'line'], hidden))).toBe('area-groups');
});

test('a visible or reversed secondary category axis stays outside the shared painters', () => {
  // A visible axis has its own label row; a reversed deleted one reverses its
  // series (Excel), which one shared category scale cannot draw.
  expect(classicPlotDispatch(combo(['bar', 'line'], { hidden: false }))).toBe('unsupported');
  expect(classicPlotDispatch(combo(['line', 'line'], { hidden: false }))).toBe('unsupported');
  const reversed = { hidden: true, orientation: 'maxMin' };
  expect(classicPlotDispatch(combo(['bar', 'line'], reversed))).toBe('unsupported');
  expect(classicPlotDispatch(combo(['bar', 'bar'], reversed))).toBe('unsupported');
  expect(classicPlotDispatch(combo(['area', 'line'], reversed))).toBe('unsupported');
  // Both reversed: a crosses="max" secondary value axis would sit on the left.
  expect(classicPlotDispatch({ ...combo(['line', 'line'], reversed), catAxisOrientation: 'maxMin' }))
    .toBe('unsupported');
});

test('date axes, crossBetween and horizontal bars keep a deleted secondary category axis outside', () => {
  expect(classicPlotDispatch(combo(['line', 'line'], { hidden: true, dateAxis: true }))).toBe('unsupported');
  expect(classicPlotDispatch({ ...combo(['line', 'line'], { hidden: true }), catAxisIsDate: true }))
    .toBe('unsupported');
  for (const kinds of [['line', 'line'], ['bar', 'bar'], ['area', 'bar']] as Array<[Kind, Kind]>) {
    const midCat = combo(kinds, { hidden: true });
    (midCat.secondaryValAxis as { crossBetween?: string }).crossBetween = 'midCat';
    expect(classicPlotDispatch(midCat)).toBe('unsupported');
  }
  const midCat = combo(['line', 'line'], { hidden: true });
  (midCat.secondaryValAxis as { crossBetween?: string }).crossBetween = 'midCat';
  expect(classicPlotDispatch({ ...midCat, catAxisCrossBetween: 'midCat' })).toBe('line-groups');
  const horizontal = combo(['bar', 'line'], { hidden: true });
  (horizontal.plotGroups as ChartPlotGroup[])[0].barDirection = 'bar';
  expect(classicPlotDispatch(horizontal)).toBe('unsupported');
});

test('the area painter keeps stacked secondary line groups outside', () => {
  // It draws line groups from raw values, so a stacked line would lose its sum.
  const model = combo(['area', 'line'], { hidden: true });
  (model.plotGroups as ChartPlotGroup[])[1].grouping = 'stacked';
  expect(classicPlotDispatch(model)).toBe('unsupported');
});
