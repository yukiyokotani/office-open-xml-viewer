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
    categories: ['a', 'b', 'c'],
    series: [{ values: [1, 2, 3] }, { values: [4, 5, 6] }],
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
  expect(classicPlotDispatch(combo(['area', 'line'], reversed))).toBe('unsupported');
  // Both reversed: a crosses="max" secondary value axis would sit on the left.
  expect(classicPlotDispatch({ ...combo(['line', 'line'], reversed), catAxisOrientation: 'maxMin' }))
    .toBe('unsupported');
});

test('date axes, crossBetween and horizontal bars keep a deleted secondary category axis outside', () => {
  expect(classicPlotDispatch(combo(['line', 'line'], { hidden: true, dateAxis: true }))).toBe('unsupported');
  expect(classicPlotDispatch({ ...combo(['line', 'line'], { hidden: true }), catAxisIsDate: true }))
    .toBe('unsupported');
  for (const kinds of [['line', 'line'], ['area', 'bar']] as Array<[Kind, Kind]>) {
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

test('a scene of bar groups keeps the bar route whether or not it shares', () => {
  // The bar painter drew every all-bar scene before sharing existed; a
  // placeholder there would lose a chart that drew.
  const variants: ChartModel[] = [
    combo(['bar', 'bar'], { hidden: true, orientation: 'maxMin' }),
    { ...combo(['bar', 'bar'], { hidden: true, orientation: 'maxMin' }), catAxisOrientation: 'maxMin' },
    combo(['bar', 'bar'], { hidden: true, dateAxis: true }),
    { ...combo(['bar', 'bar'], { hidden: true }), catAxisIsDate: true },
  ];
  const midCat = combo(['bar', 'bar'], { hidden: true });
  (midCat.secondaryValAxis as { crossBetween?: string }).crossBetween = 'midCat';
  const horizontal = combo(['bar', 'bar'], { hidden: true });
  for (const group of horizontal.plotGroups as ChartPlotGroup[]) group.barDirection = 'bar';
  const longer = combo(['bar', 'bar'], { hidden: true });
  longer.series[1] = { ...longer.series[1], categories: ['a', 'b', 'c', 'd'], values: [1, 2, 3, 4] };
  for (const model of [...variants, midCat, horizontal, longer]) {
    expect(classicPlotDispatch(model)).toBe('bar-combo');
  }
});

test('a series without its own categories carries the chart category count', () => {
  // The parser leaves the first series' categories on the chart, so a first
  // series with fewer values than categories still shares them.
  const short = combo(['bar', 'line'], { hidden: true });
  short.series[0] = { ...short.series[0], values: [4] };
  short.series[1] = { ...short.series[1], categories: ['a', 'b', 'c'] };
  expect(classicPlotDispatch(short)).toBe('bar-combo');
  const lines = combo(['line', 'line'], { hidden: true });
  lines.series[0] = { ...lines.series[0], values: [4] };
  expect(classicPlotDispatch(lines)).toBe('line-groups');
});

test('a deleted secondary axis with its own category count stays outside', () => {
  // Its categories spread across the plot independently; the shared scale
  // would truncate or misplace them.
  const longer = combo(['line', 'line'], { hidden: true });
  longer.series[1] = { ...longer.series[1], categories: ['a', 'b', 'c', 'd', 'e'], values: [1, 2, 3, 4, 5] };
  expect(classicPlotDispatch(longer)).toBe('unsupported');
  const shorter = combo(['bar', 'line'], { hidden: true });
  shorter.series[1] = { ...shorter.series[1], categories: ['a', 'b'], values: [1, 2] };
  expect(classicPlotDispatch(shorter)).toBe('unsupported');
  // A longer primary group behind a secondary group listed first.
  const primaryLonger = combo(['line', 'line'], { hidden: true });
  primaryLonger.series[0] = { ...primaryLonger.series[0], categories: ['a', 'b', 'c', 'd', 'e'], values: [1, 2, 3, 400, 500] };
  expect(classicPlotDispatch(primaryLonger)).toBe('unsupported');
  // Blank trailing values over the same categories still share them.
  const blanks = combo(['bar', 'line'], { hidden: true });
  blanks.series[1] = { ...blanks.series[1], values: [1, null, null] };
  expect(classicPlotDispatch(blanks)).toBe('bar-combo');
});

test('the area painter keeps stacked secondary line groups outside', () => {
  // It draws line groups from raw values, so a stacked line would lose its sum.
  const model = combo(['area', 'line'], { hidden: true });
  (model.plotGroups as ChartPlotGroup[])[1].grouping = 'stacked';
  expect(classicPlotDispatch(model)).toBe('unsupported');
});
