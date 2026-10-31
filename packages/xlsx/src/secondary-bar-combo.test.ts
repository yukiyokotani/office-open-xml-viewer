/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, expect, it } from 'vitest';
import { renderChart, type ChartModel } from '@silurus/ooxml-core';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { secondaryBarComboXlsxBytes } from '../../../tests/fixtures/secondary-bar-combo-xlsx.mjs';
import { initSync, XlsxArchive } from './wasm/xlsx_parser.js';
import type { Worksheet } from './types.js';

beforeAll(() => {
  initSync({ module: new Uint8Array(readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url))) });
});

function parsedChart(bytes: Uint8Array): ChartModel {
  const archive = new XlsxArchive(bytes, undefined, undefined, undefined);
  const decode = (data: Uint8Array) => JSON.parse(new TextDecoder().decode(data));
  try {
    decode(archive.parse());
    archive.open_sheet_cursor(0, 'Combo');
    for (let pull = 0; pull < 4; pull++) {
      const unit = decode(archive.pull_sheet_cursor(128)) as
        | { kind: 'preview' } | { kind: 'rows' } | { kind: 'finished'; worksheet: Worksheet };
      if (unit.kind === 'finished') return unit.worksheet.charts[0].chart;
    }
    throw new Error('worksheet did not finish');
  } finally {
    archive.free();
  }
}

/** Records the text and filled rectangles a chart paints. */
function paint(chart: ChartModel): { texts: string[]; rects: number } {
  const texts: string[] = [];
  let rects = 0;
  const state: Record<string, unknown> = { font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000',
    lineWidth: 1, textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1 };
  const gradient = { addColorStop() {} };
  const ctx = new Proxy(state, {
    get(target, prop: string) {
      if (prop in target) return target[prop];
      if (prop === 'measureText') return (text: string) => ({ width: String(text).length * 6 });
      if (prop === 'fillText') return (text: string) => { texts.push(String(text)); };
      if (prop === 'fillRect') return () => { rects++; };
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => gradient;
      if (prop === 'getLineDash') return () => [];
      return () => undefined;
    },
    set(target, prop: string, value) { target[prop] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderChart(ctx, chart, { x: 0, y: 0, w: 480, h: 288 }, 1);
  return { texts, rects };
}

// Two column groups on primary and secondary value axes over a deleted
// secondary category axis: the bar painter draws them whether or not the
// groups can share one category scale.
it.each([
  ['shared categories, deleted axis at b', {}],
  ['shared categories, deleted axis at t', { deletedAxPos: 't' }],
  ['a date primary category axis', { dateAxis: true }],
  ['a date primary category axis, deleted axis at t', { dateAxis: true, deletedAxPos: 't' }],
  ['midCat on the secondary value axis, deleted axis at t', { midCat: true, deletedAxPos: 't' }],
  ['reversed category axes, deleted axis at t', { reversed: true, deletedAxPos: 't' }],
  ['a first series shorter than its categories, deleted axis at t', { shortFirst: true, deletedAxPos: 't' }],
])('draws a secondary-axis column combo with %s', (_name, variant) => {
  const chart = parsedChart(secondaryBarComboXlsxBytes(variant));
  expect(chart.plotGroups?.map(group => group.kind)).toEqual(['bar', 'bar']);
  const painted = paint(chart);
  expect(painted.texts).not.toContain('Unsupported chart');
  // Background plus at least one bar per non-empty point.
  expect(painted.rects).toBeGreaterThanOrEqual(5);
});
