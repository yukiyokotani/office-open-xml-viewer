import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries, ChartSeriesDataLabels } from '../types/chart';
import { renderChart } from './renderer.js';
import { renderChartExChart } from './chart-ex-renderer.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';
import { chartPaintTexts } from '../fonts/chart-paint-texts.js';

// Scope guard for #2346 on chart category text (chart-number-format.ts
// `localizedCategoryText`): separators only. For one representative chart per
// paint-site class,
// - en-US paints exactly the labels listed here, which are what the renderer
//   painted before `categoriesNumeric` existed (the model without the flag
//   paints the same list, in the same places);
// - fr-FR paints the same labels in the same places, differing only in the
//   decimal (`,`) and group (U+202F) separators of numbers;
// - a text category that merely looks like a number is not repainted;
// - font demand (`chartPaintTexts`) covers every character painted.


interface Label { text: string; x: number; y: number; align: string; baseline: string; font: string }
interface Segment { x1: number; y1: number; x2: number; y2: number }

interface FilledRect { x: number; y: number; w: number; h: number }
interface ArcCall { r: number }

function recordingContext(): {
  ctx: CanvasRenderingContext2D; labels: Label[]; segments: Segment[]; rects: FilledRect[]; arcs: ArcCall[];
} {
  const rects: FilledRect[] = [];
  const arcs: ArcCall[] = [];
  const labels: Label[] = [];
  const segments: Segment[] = [];
  const path: Array<{ x: number; y: number }> = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1, lineCap: 'butt', lineJoin: 'miter',
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (text: string) => ({ width: Array.from(String(text)).length * 7 });
        case 'beginPath': return () => { path.length = 0; };
        case 'moveTo': return (x: number, y: number) => { path.length = 0; path.push({ x, y }); };
        case 'lineTo': return (x: number, y: number) => { path.push({ x, y }); };
        case 'stroke': return () => {
          if (path.length === 2) segments.push({ x1: path[0].x, y1: path[0].y, x2: path[1].x, y2: path[1].y });
        };
        case 'fillText': return (text: string, x: number, y: number) => labels.push({
          text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline), font: String(state.font),
        });
        case 'getLineDash': return () => [];
        case 'fillRect': return (x: number, y: number, w: number, h: number) => { rects.push({ x, y, w, h }); };
        case 'rect': return (x: number, y: number, w: number, h: number) => { rects.push({ x, y, w, h }); };
        case 'arc': return (_x: number, _y: number, r: number) => { arcs.push({ r }); };
        case 'ellipse': return (_x: number, _y: number, rx: number, ry: number) => { arcs.push({ r: Math.min(rx, ry) }); };
        case 'createLinearGradient':
        case 'createRadialGradient': return () => ({ addColorStop() {} });
        case 'createPattern': return () => null;
        default: return () => undefined;
      }
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, labels, segments, rects, arcs };
}

function series(over: Partial<ChartSeries> = {}): ChartSeries {
  return { name: 'S1', color: '4472C4', values: [10, 20], ...over };
}

function model(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredColumn',
    title: null,
    categories: ['Alpha', 'Beta'],
    series: [series()],
    showDataLabels: false,
    valMin: 0,
    valMax: 20,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: true,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'none',
    catAxisMajorTickMark: 'none',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: 1000,
    valAxisFontSizeHpt: 1000,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisMajorGridlines: false,
    ...over,
  };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

function render(chart: ChartModel): ReturnType<typeof recordingContext> {
  const rec = recordingContext();
  renderChart(
    rec.ctx, chart, RECT, 1, 0,
    { render: renderSimpleThreeDChart }, undefined, undefined,
    { render: renderChartExChart },
  );
  return rec;
}

afterEach(() => { vi.unstubAllGlobals(); });


function inLocale<T>(locale: string, run: () => T): T {
  vi.stubGlobal('navigator', { language: locale });
  try { return run(); } finally { vi.unstubAllGlobals(); }
}

/** fr-FR text with its separators mapped back to the invariant pair. */
const invariant = (text: string): string => text.replace(/,/g, '.').replace(/\u202f/g, ',');

/** The model with every `categoriesNumeric` removed: the pre-#2346 input. */
function withoutProvenance(chart: ChartModel): ChartModel {
  const { categoriesNumeric: _chart, ...rest } = chart;
  return {
    ...rest,
    series: chart.series.map(({ categoriesNumeric: _series, ...series }) => series),
  };
}

const labelled: ChartSeriesDataLabels = { showVal: false, showCatName: true, showSerName: false, showPercent: false };
const secondaryAxis = { min: 0, max: 20, title: null, hidden: false, lineHidden: true, majorTickMark: 'none', fontSizeHpt: 1000 } as never;
const table = { showHorzBorder: true, showVertBorder: true, showOutline: true, showKeys: false } as never;

interface Site {
  name: string;
  chart: ChartModel;
  /** Labels en-US paints that come from a category (a subset of all labels). */
  en: string[];
  /** The same labels under fr-FR. */
  fr: string[];
  /** The labels go through a number format, so provenance plays no part. */
  formatted?: boolean;
}

const sites: Site[] = [
  {
    name: 'column axis',
    chart: model({ categories: ['1.2', '1234.5'], categoriesNumeric: true }),
    en: ['1.2', '1234.5'],
    fr: ['1,2', '1234,5'],
  },
  {
    name: 'column axis under a grouping code',
    formatted: true,
    chart: model({ categories: ['1.2', '1234.5'], categoriesNumeric: true, catAxisFormatCode: '#,##0.00' }),
    en: ['1.20', '1,234.50'],
    fr: ['1,20', '1\u202f234,50'],
  },
  {
    name: 'line point label (showCatName) and its axis',
    chart: model({
      chartType: 'line', categories: ['1.2', '2.5'], categoriesNumeric: true,
      series: [series({ seriesDataLabels: labelled })],
    }),
    en: ['1.2', '2.5', '1.2', '2.5'],
    fr: ['1,2', '2,5', '1,2', '2,5'],
  },
  {
    // The first series' cache code (0%) is not a legend format: the legend
    // paints the raw category.
    name: 'pie legend and slice labels',
    chart: model({
      chartType: 'pie', categories: ['0.25', '0.75'], categoriesNumeric: true, showLegend: true, legendPos: 'r',
      series: [series({ seriesDataLabels: labelled, catFormatCode: '0%' })],
    }),
    en: ['0.25', '0.75', '0.25', '0.75'],
    fr: ['0,25', '0,75', '0,25', '0,75'],
  },
  {
    // A scatter X label is the X value formatted as a number (General here),
    // whatever the source kind: the formatter paints the separator.
    name: 'scatter X label, per-series X lists',
    formatted: true,
    chart: model({
      chartType: 'scatter', categories: ['1.25', '2.75'], categoriesNumeric: true, valMax: 40,
      series: [
        series({ seriesType: 'scatter', seriesDataLabels: labelled }),
        series({ seriesType: 'scatter', name: 'S2', values: [30, 40], categories: ['3.25', '4.75'], categoriesNumeric: true, seriesDataLabels: labelled }),
      ],
    }),
    en: ['1.25', '2.75', '3.25', '4.75'],
    fr: ['1,25', '2,75', '3,25', '4,75'],
  },
  {
    name: 'data table header',
    chart: model({ dataTable: table, categories: ['1.2', '2.5'], categoriesNumeric: true }),
    en: ['1.2', '2.5'],
    fr: ['1,2', '2,5'],
  },
  {
    name: 'secondary category axis from the secondary series\' own list',
    chart: model({
      categories: ['Alpha'],
      secondaryValAxis: secondaryAxis, secondaryCatAxis: secondaryAxis,
      series: [series({ values: [1] }), series({ values: [2], useSecondaryAxis: true, categories: ['1.2'], categoriesNumeric: true })],
    }),
    en: ['1.2'],
    fr: ['1,2'],
  },
  {
    name: '3-D column axis',
    chart: model({
      chartType: 'clusteredBar',
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      plotGroups: [{ kind: 'bar3D', seriesStart: 0, seriesCount: 1, categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none', varyColors: false, barDirection: 'col', grouping: 'clustered' }],
      categories: ['1.2', '2.5'], categoriesNumeric: true,
    }),
    en: ['1.2', '2.5'],
    fr: ['1,2', '2,5'],
  },
];

describe('chart category text: separators only (#2346)', () => {
  for (const site of sites) {
    it(site.name, () => {
      const en = inLocale('en-US', () => render(site.chart)).labels;
      const before = inLocale('en-US', () => render(withoutProvenance(site.chart))).labels;
      const fr = inLocale('fr-FR', () => render(site.chart)).labels;
      // en-US: the provenance flag changes nothing, text or placement.
      expect(en).toEqual(before);
      const categoryLabels = new Set(site.en);
      expect(en.map(entry => entry.text).filter(text => categoryLabels.has(text))).toEqual(site.en);
      // fr-FR: the same labels in the same places; only separators differ.
      expect(fr.map(entry => ({ ...entry, text: invariant(entry.text) }))).toEqual(en);
      const localized = new Set(site.fr);
      expect(fr.map(entry => entry.text).filter(text => localized.has(text))).toEqual(site.fr);
    });
  }

  it('font demand covers every character the sites paint under fr-FR', () => {
    for (const site of sites) {
      const demanded = inLocale('fr-FR', () => chartPaintTexts(site.chart));
      expect(demanded, site.name).not.toBe('all');
      const covered = new Set(Array.from((demanded as readonly string[]).join('')));
      const painted = inLocale('fr-FR', () => render(site.chart)).labels.flatMap(entry => Array.from(entry.text));
      expect(painted.filter(char => !covered.has(char)), site.name).toEqual([]);
    }
  });

  it('a text category that reads like a number is not repainted', () => {
    for (const site of sites) {
      if (site.formatted) continue;
      const text = withoutProvenance(site.chart);
      const fr = inLocale('fr-FR', () => render(text)).labels.map(entry => entry.text);
      const categoryLabels = new Set(site.en);
      expect(fr.filter(label => categoryLabels.has(label)), site.name).toEqual(site.en);
    }
  });
});
