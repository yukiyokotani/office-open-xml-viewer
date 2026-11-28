import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries, ChartSeriesDataLabels } from '../types/chart';
import { renderChart } from './renderer.js';
import { renderChartExChart } from './chart-ex-renderer.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';
import { categorySourceAt } from './category-caption.js';
import { wrapMeasuredText } from './shared/axis.js';
import { formatLocalizedExcelShortDate } from './chart-number-format.js';
import { adoptHostLocale } from '../number-symbols.js';

// Category text keeps its numeric-vs-text source (`categoriesNumeric`, per
// series for a series' own categories) at every paint site: a numeric
// category is a number (General: host decimal separator), a text category is
// never reformatted. Each case runs under fr-FR and en-US.

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
// 10pt at ptToPx 1 → 10 px font; stacked axis lines are font + 2 px apart.
const LINE = 12;

function render(chart: ChartModel): ReturnType<typeof recordingContext> {
  const rec = recordingContext();
  renderChart(
    rec.ctx, chart, RECT, 1, 0,
    { render: renderSimpleThreeDChart }, undefined, undefined,
    { render: renderChartExChart },
  );
  return rec;
}

function label(rec: ReturnType<typeof recordingContext>, text: string): Label {
  const found = rec.labels.find(entry => entry.text === text);
  if (!found) throw new Error(`no fillText(${JSON.stringify(text)}) in ${JSON.stringify(rec.labels.map(entry => entry.text))}`);
  return found;
}

function expectNoJoinedText(rec: ReturnType<typeof recordingContext>): void {
  expect(rec.labels.filter(entry => /[\r\n]/.test(entry.text))).toEqual([]);
}


afterEach(() => { vi.unstubAllGlobals(); });

function inLocale<T>(locale: string, run: () => T): T {
  vi.stubGlobal('navigator', { language: locale });
  try { return run(); } finally { vi.unstubAllGlobals(); }
}
const texts = (rec: ReturnType<typeof recordingContext>) => rec.labels.map(entry => entry.text);
const labelled: ChartSeriesDataLabels = { showVal: false, showCatName: true, showSerName: false, showPercent: false };

describe('single-series line chart, category 1.2', () => {
  const line = (numeric: boolean) => model({
    chartType: 'line', categories: ['1.2', '2.5'], categoriesNumeric: numeric,
    catAxisFormatCode: numeric ? undefined : '0.0',
    series: [series({ seriesDataLabels: labelled })],
  });

  it('numeric: axis and point label agree on the host decimal separator', () => {
    const fr = texts(inLocale('fr-FR', () => render(line(true))));
    expect(fr.filter(text => text === '1,2')).toHaveLength(2);
    expect(fr).not.toContain('1.2');
    const en = texts(inLocale('en-US', () => render(line(true))));
    expect(en.filter(text => text === '1.2')).toHaveLength(2);
  });

  it('text with identical authored characters stays verbatim, axis code or not', () => {
    for (const locale of ['fr-FR', 'en-US']) {
      const painted = texts(inLocale(locale, () => render(line(false))));
      expect(painted.filter(text => text === '1.2'), locale).toHaveLength(2);
      expect(painted, locale).not.toContain('1,2');
    }
  });

  it('numeric pie legend and data labels use the same formatter', () => {
    const pie = (numeric: boolean) => model({
      chartType: 'pie', categories: ['1.2', '2.5'], categoriesNumeric: numeric, showLegend: true, legendPos: 'r',
      series: [series({ seriesDataLabels: labelled })],
    });
    const numericPie = texts(inLocale('fr-FR', () => render(pie(true))));
    expect(numericPie.filter(text => text === '1,2').length).toBeGreaterThanOrEqual(2);
    const textPie = texts(inLocale('fr-FR', () => render(pie(false))));
    expect(textPie).not.toContain('1,2');
    expect(textPie).toContain('1.2');
  });
});

describe('categories follow the source that supplies them', () => {
  const secondaryAxis = { min: 0, max: 20, title: null, hidden: false, lineHidden: true, majorTickMark: 'none', fontSizeHpt: 1000 } as never;
  // Primary categories 5.5; the secondary series' own cache says 1.2.
  const combo = (primaryNumeric: boolean, secondaryNumeric: boolean) => model({
    categories: ['5.5'], categoriesNumeric: primaryNumeric,
    secondaryValAxis: secondaryAxis, secondaryCatAxis: secondaryAxis,
    series: [series({ values: [1] }), series({ values: [2], useSecondaryAxis: true, categories: ['1.2'], categoriesNumeric: secondaryNumeric })],
  });

  it('secondary axis through the bar renderer: text primary + numeric secondary', () => {
    const painted = texts(inLocale('fr-FR', () => render(combo(false, true))));
    expect(painted).toContain('1,2');
    expect(painted).toContain('5.5');
  });

  it('secondary axis through the bar renderer: numeric primary + text secondary', () => {
    const painted = texts(inLocale('fr-FR', () => render(combo(true, false))));
    expect(painted).toContain('1.2');
    expect(painted).not.toContain('1,2');
    expect(painted).toContain('5,5');
  });

  it('resolves per point: a short own numeric cache falls back to the shared text list', () => {
    const chart = model({
      categories: ['a', '7.5'], categoriesNumeric: false,
      series: [series({ categories: ['1.2'], categoriesNumeric: true })],
    });
    expect(categorySourceAt(chart, chart.series[0], 0)).toMatchObject({ text: '1.2', numeric: true, owner: 'series' });
    expect(categorySourceAt(chart, chart.series[0], 1)).toMatchObject({ text: '7.5', numeric: false, owner: 'shared' });
    expect(categorySourceAt(chart, null, 1)).toMatchObject({ text: '7.5', numeric: false, owner: 'shared' });
  });

  it('column chart: the axis follows the axis code, the point label its own source', () => {
    const column = (over: Partial<ChartSeries>) => model({
      categories: ['1.2', '2.5'], categoriesNumeric: true, catAxisFormatCode: '0.00',
      series: [series({
        seriesDataLabels: { ...labelled, showVal: false },
        dataLabelOverrides: [{ idx: 0, showCatName: true } as never, { idx: 1, showCatName: true } as never],
        ...over,
      })],
    });
    // Axis ticks '1,20'; the point label is General ('1,2'): only the label can
    // paint that text, the axis cannot satisfy it. No axis-code inheritance.
    const general = texts(inLocale('fr-FR', () => render(column({}))));
    expect(general).toContain('1,20');
    expect(general).toContain('1,2');
    // A source code of its own formats the label and not the axis.
    const own = texts(inLocale('fr-FR', () => render(model({
      categories: ['1.2', '2.5'], categoriesNumeric: true, catAxisFormatCode: '0.00', categoryFormatCode: '0.000',
      series: [series({
        seriesDataLabels: { ...labelled, showVal: false },
        dataLabelOverrides: [{ idx: 0, showCatName: true } as never],
      })],
    }))));
    expect(own).toContain('1,20');
    expect(own).toContain('1,200');
  });

  it('a short own cache falls back to the shared source with the shared source\'s own format', () => {
    const chart = model({
      categories: ['a', '5.6'], categoriesNumeric: false, categoryFormatCode: '0.0',
      series: [series({ categories: ['1.2'], categoriesNumeric: true, catFormatCode: '0.000' })],
    });
    expect(categorySourceAt(chart, chart.series[0], 0)).toMatchObject({ text: '1.2', formatCode: '0.000', owner: 'series' });
    expect(categorySourceAt(chart, chart.series[0], 1)).toMatchObject({ text: '5.6', numeric: false, formatCode: '0.0', owner: 'shared' });
  });

  it('generated captions are resolved by index, not by rebuilding the 1..n list per label', () => {
    let valueReads = 0;
    const base = series({ values: Array.from({ length: 4000 }, (_, i) => i) });
    const counted = { ...base };
    Object.defineProperty(counted, 'values', { get() { valueReads++; return base.values; } });
    const chart = model({ categories: [], series: [counted] });
    for (let i = 0; i < 4000; i++) categorySourceAt(chart, null, i);
    expect(categorySourceAt(chart, null, 3999).text).toBe('4000');
    expect(categorySourceAt(chart, null, 4000).text).toBe('');
    expect(valueReads).toBeLessThanOrEqual(2);
  });

  it('with no categories at all point labels keep the generated 1..n names; pie and legend do not', () => {
    const chart = model({ categories: [], series: [series({ values: [5, 6] })] });
    expect(categorySourceAt(chart, chart.series[0], 1, 'chart').text).toBe('2');
    expect(categorySourceAt(chart, chart.series[0], 1, 'none')).toMatchObject({ text: '', owner: 'generated' });
    expect(categorySourceAt(chart, chart.series[0], 1, 'index').text).toBe('2');
  });

  it('scatter: a text-backed X caption stays verbatim although the coordinate parses as a number', () => {
    const scatter = (numeric: boolean) => model({
      chartType: 'scatter', categories: ['1.2', '2.5'], categoriesNumeric: numeric,
      series: [series({ seriesType: 'scatter', categories: ['1.2', '2.5'], categoriesNumeric: numeric, seriesDataLabels: labelled })],
    });
    const text = texts(inLocale('fr-FR', () => render(scatter(false))));
    expect(text).toContain('1.2');
    expect(text).not.toContain('1,2');
    const numeric = texts(inLocale('fr-FR', () => render(scatter(true))));
    expect(numeric).toContain('1,2');
  });

  it('plotVisibleOnly keeps the shared source provenance on scatter series it filters', () => {
    const scatter = (plotVisibleOnly: boolean) => model({
      chartType: 'scatter', categories: ['1.2', '2.5', '3.5'], categoriesNumeric: false, plotVisibleOnly,
      series: [series({ seriesType: 'scatter', values: [1, 2, 3], sourceHidden: [false, true, false], seriesDataLabels: labelled })],
    });
    for (const visibleOnly of [false, true]) {
      const painted = texts(inLocale('fr-FR', () => render(scatter(visibleOnly))));
      expect(painted, `visibleOnly=${visibleOnly}`).toContain('1.2');
      expect(painted, `visibleOnly=${visibleOnly}`).not.toContain('1,2');
    }
  });

  it('a pie legend resolves entries beyond a short own cache from the shared source', () => {
    const pie = model({
      chartType: 'pie', categories: ['a', '5.6'], categoriesNumeric: true, showLegend: true, legendPos: 'r',
      series: [series({ values: [1, 2], categories: ['1.2'], categoriesNumeric: true, seriesDataLabels: { ...labelled, showCatName: true } })],
    });
    const painted = texts(inLocale('fr-FR', () => render(pie)));
    expect(painted).toContain('5,6');
    expect(painted).not.toContain('Item 2');
  });

  it('surface: the authored 0.00 axis code formats a numeric category', () => {
    const surface = model({
      chartType: 'surface', categories: ['1.2', '2.5', '3'], categoriesNumeric: true, catAxisFormatCode: '0.00',
      valMin: 0, valMax: 10, valAxisMajorUnit: 5,
      series: [series({ values: [1, 9, 2] }), series({ values: [9, 1, 8] })],
    });
    const painted = texts(inLocale('fr-FR', () => render(surface)));
    expect(painted).toContain('1,20');
    expect(painted).not.toContain('1.2');
  });

  const threeD = (over: Partial<ChartModel>) => model({
    chartType: 'clusteredBar',
    threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
    plotGroups: [{ kind: 'bar3D', seriesStart: 0, seriesCount: 2, categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none', varyColors: false, barDirection: 'col', grouping: 'clustered' }],
    ...over,
  });

  it('3-D axis: the first series with its own categories decides, both directions', () => {
    const textChartNumericSeries = threeD({
      categories: ['1.2'], categoriesNumeric: false,
      series: [series({ values: [1] }), series({ values: [2], categories: ['1.2'], categoriesNumeric: true })],
    });
    const painted = texts(inLocale('fr-FR', () => render(textChartNumericSeries)));
    expect(painted).toContain('1,2');
    const numericChartTextSeries = threeD({
      categories: ['1.2'], categoriesNumeric: true, catAxisFormatCode: '0.0',
      series: [series({ values: [1] }), series({ values: [2], categories: ['1.2'], categoriesNumeric: false })],
    });
    const verbatim = texts(inLocale('fr-FR', () => render(numericChartTextSeries)));
    expect(verbatim).toContain('1.2');
    expect(verbatim).not.toContain('1,2');
  });
});

describe('data table header provenance', () => {
  const table = { showHorzBorder: true, showVertBorder: true, showOutline: true, showKeys: false } as never;
  const withTable = (over: Partial<ChartModel>) => model({
    chartType: 'clusteredColumn', dataTable: table, categories: ['1.2'], ...over,
  });

  it('a text header "1.2" is never read as built-in date format 14 from another series', () => {
    const textHeader = withTable({
      categoriesNumeric: false,
      series: [series({ values: [1] }), series({ values: [2], catFormatBuiltinId: 14 })],
    });
    const painted = texts(inLocale('fr-FR', () => render(textHeader)));
    expect(painted).toContain('1.2');
    expect(painted).not.toContain(formatLocalizedExcelShortDate(1.2, false, 'fr-FR'));
  });

  it('the header format comes from the series supplying the headers', () => {
    const header = withTable({
      categoriesNumeric: true, catAxisFormatCode: '0.00',
      series: [series({ values: [1] }), series({ values: [2], catFormatCode: '0.000' })],
    });
    const painted = texts(inLocale('fr-FR', () => render(header)));
    expect(painted).toContain('1,20');
    expect(painted).not.toContain('1,200');
  });
});

describe('built-in short dates follow the adopted host locale', () => {
  it('formatLocalizedExcelShortDate defaults to the adopted locale, not navigator', () => {
    vi.stubGlobal('navigator', { language: 'en-US' });
    expect(formatLocalizedExcelShortDate(45_659, false)).toBe('1/2/2025');
    adoptHostLocale('fr-FR');
    expect(formatLocalizedExcelShortDate(45_659, false)).toBe('02/01/2025');
  });
});

describe('label wrapping keeps no-break group separators', () => {
  const ctx = { measureText: (text: string) => ({ width: Array.from(text).length * 7 }) } as unknown as CanvasRenderingContext2D;

  it('does not break at, or rewrite, U+202F / U+00A0 / U+2007', () => {
    for (const space of ['\u202f', '\u00a0', '\u2007']) {
      expect(wrapMeasuredText(ctx, `1${space}234,5`, 100)).toEqual([`1${space}234,5`]);
    }
    // An ordinary space still wraps.
    expect(wrapMeasuredText(ctx, 'ab cd', 20)).toEqual(['ab', 'cd']);
  });

  it('the emergency code-point split never breaks beside a no-break space', () => {
    // 7 px per glyph, 14 px wide: the whole token cannot fit, so it is split by code point.
    for (const space of ['\u202f', '\u00a0', '\u2007']) {
      const lines = wrapMeasuredText(ctx, `1${space}234,5`, 14);
      expect(lines.join(''), space).toBe(`1${space}234,5`);
      for (const line of lines) {
        expect(line.startsWith(space) || line.endsWith(space), `${JSON.stringify(lines)}`).toBe(false);
      }
    }
  });

  it('a fr-FR numeric category paints its U+202F group separator through the axis', () => {
    const painted = texts(inLocale('fr-FR', () => render(model({
      categories: ['1234.5', '2'], categoriesNumeric: true, catAxisFormatCode: '#,##0.0',
    }))));
    expect(painted).toContain('1\u202f234,5');
  });
});
