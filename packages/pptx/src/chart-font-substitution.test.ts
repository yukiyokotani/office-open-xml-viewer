import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { renderSlideWithEmbeddedFonts } from './renderer.js';
import { PptxFontPreloadAccumulator } from './font-requests.js';
import type { Slide } from './types';
import { officeFaceWeight } from './powerpoint-line-metrics.js';

// Charts demand the faces they name and paint through the same loaded
// substitution map as shape text (LoadOptions.fonts).

function chart(overrides: Partial<ChartModel> = {}): ChartModel {
  return {
    chartType: 'clusteredBar', title: 'Chart Title', categories: ['A', 'B'],
    series: [{ name: 'Series', color: null, values: [1, 2] }],
    showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null, valAxisTitle: null,
    catAxisHidden: false, valAxisHidden: false, catAxisLineHidden: false, valAxisLineHidden: false,
    plotAreaBg: null, chartBg: null, showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out', catAxisMajorTickMark: 'out', titleFontSizeHpt: null, titleFontColor: null,
    titleFontFace: 'Chart Face', catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
    ...overrides,
  } as ChartModel;
}

const slide = (model: ChartModel): Slide => ({
  index: 0, slideNumber: 1, background: null,
  elements: [{ type: 'chart', x: 0, y: 0, width: 4_000_000, height: 3_000_000, rotation: 0, flipH: false, flipV: false, chart: model }],
} as unknown as Slide);

async function paintedFonts(target: Slide, substitutions?: ReadonlyMap<string, { family: string; script?: 'arabic' }>) {
  const calls: Array<{ text: string; font: string }> = [];
  let font = '10px sans-serif';
  const noop = () => {};
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: text.length * 7, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
    fillText: (text: string) => { calls.push({ text, font }); },
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    getLineDash: () => [],
  } as Record<string | symbol, unknown>, {
    get(target, key) { return key in target ? target[key] : noop; },
    set(target, key, value) { target[key] = value; return true; },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;
  await renderSlideWithEmbeddedFonts(canvas, target, 9_144_000, 6_858_000, {
    width: 960, dpr: 1, majorFont: null, minorFont: null, defaultTextColor: null,
    ...(substitutions ? { fontSources: { substitutions, families: new Map([['served face', 'ooxml-font-source-1-0']]) } } : {}),
  });
  return calls;
}

describe('chart text and font sources', () => {
  it('demands the faces a chart names, resolving theme references', () => {
    const accumulator = new PptxFontPreloadAccumulator(null, null, undefined, undefined, undefined, undefined, officeFaceWeight);
    accumulator.addSlide(slide(chart({
      legendFontFace: '+mn-lt', themeMinorFontLatin: 'Chart Body', catAxisFontFace: 'Axis Face',
    })));
    const names = accumulator.names();
    const facts = accumulator.demandFacts()!;
    expect(names.slice(0, facts.familyCount)).toEqual(expect.arrayContaining(['Chart Face', 'Chart Body', 'Axis Face']));
    // Without font sources the preflight names are unchanged.
    const plain = new PptxFontPreloadAccumulator(null, null);
    plain.addSlide(slide(chart()));
    expect(plain.names()).not.toContain('Chart Face');
  });

  it('paints chart text through the private family of the loaded unscoped substitute', async () => {
    const offline = await paintedFonts(slide(chart()));
    const title = (calls: typeof offline) => calls.find((call) => call.text === 'Chart Title')?.font ?? '';
    expect(title(offline)).toContain('"Chart Face"');
    expect(title(offline)).not.toContain('ooxml-font-source');
    const loaded = await paintedFonts(slide(chart()), new Map([['chart face', { family: 'Served Face' }]]));
    expect(title(loaded)).toContain('"Chart Face", "ooxml-font-source-1-0"');
    // A script-scoped substitute never reaches chart text.
    const scoped = await paintedFonts(slide(chart()), new Map([['chart face', { family: 'Served Face', script: 'arabic' }]]));
    expect(title(scoped)).toBe(title(offline));
  });
});
