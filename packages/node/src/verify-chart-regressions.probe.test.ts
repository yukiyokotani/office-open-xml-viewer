// End-to-end verification of the two chart regressions against the REAL
// private fixtures, rendered through the actual WASM parser + core renderer on
// skia. Gated on skia + the git-ignored fixtures, so it skips cleanly where
// either is absent (CI has no private samples). This is a local acceptance
// probe, not a committed regression guard — the parse/render logic is pinned by
// the ooxml-common `parse_chart_part_*` tests and the core
// `renderer-scatter-line-and-pie-labels` unit tests.

import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadSkiaForTests, importForTests } from './test-imports';
import type { ChartModel } from '../../core/src/types/chart.ts';

const skia = await loadSkiaForTests();
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
const SAMPLE_14 = resolve(ROOT, 'packages/pptx/public/private/pptx/sample-14.pptx');
const SAMPLE_30 = resolve(ROOT, 'packages/xlsx/public/private/xlsx/sample-30.xlsx');
const SAMPLE_17 = resolve(ROOT, 'packages/pptx/public/private/pptx/sample-17.pptx');
const SAMPLE_18 = resolve(ROOT, 'packages/pptx/public/private/pptx/sample-18.pptx');

const varySamples = [SAMPLE_17, SAMPLE_18].filter((p) => existsSync(p));
const pptxMod = skia && (existsSync(SAMPLE_14) || varySamples.length > 0)
  ? await importForTests(() => import('./pptx.ts'), './pptx.ts (pptx WASM)')
  : null;
const xlsxMod = skia && existsSync(SAMPLE_30)
  ? await importForTests(() => import('./xlsx.ts'), './xlsx.ts (xlsx WASM)')
  : null;
const CORE_RENDERER = resolve(ROOT, 'packages/core/src/chart/renderer.ts');
const coreMod = skia
  ? await importForTests(() => import(CORE_RENDERER), 'packages/core/src/chart/renderer.ts')
  : null;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** Depth-first collect every `type:'chart'` element's ChartModel from a slide. */
function collectCharts(node: Any, out: ChartModel[]): void {
  if (!node || typeof node !== 'object') return;
  if (node.type === 'chart' && node.chart) out.push(node.chart as ChartModel);
  for (const k of Object.keys(node)) {
    const v = (node as Any)[k];
    if (Array.isArray(v)) v.forEach(c => collectCharts(c, out));
    else if (v && typeof v === 'object') collectCharts(v, out);
  }
}

/** Render a chart, capturing fillText calls with the fillStyle in effect, a
 *  count of series-colored (matchColor) connecting-line strokes, and the legend
 *  KEY kind (marker glyph vs. line swatch) drawn in the bottom legend band. The
 *  legend band is the strip below `LEGEND_Y_MIN`; a marker key fills a path
 *  (`fill`/`fillRect`), a line key is a 2-vertex horizontal `stroke`. Used by
 *  the scatter legend-key check (#803). */
const RECT_H = 400;
const LEGEND_Y_MIN = RECT_H - 40; // bottom legend strip (plot is above).
function renderCapture(
  chart: ChartModel,
  renderChart: (ctx: unknown, c: ChartModel, r: unknown, p?: number) => void,
  matchColor: string,
): {
  texts: Array<{ text: string; fill: string }>;
  seriesStrokes: number;
  legendKeys: Array<'marker' | 'line'>;
} {
  const canvas = new Canvas(640, RECT_H);
  const raw = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const texts: Array<{ text: string; fill: string }> = [];
  const legendKeys: Array<'marker' | 'line'> = [];
  let verts = 0;
  let lastY = 0;
  let seriesStrokes = 0;
  let pendingKey: 'marker' | 'line' | null = null;
  const inBand = (): boolean => lastY >= LEGEND_Y_MIN;
  const proxy = new Proxy(raw, {
    get(t, p: string) {
      if (p === 'fillText') {
        return (text: string, x: number, y: number) => {
          texts.push({ text, fill: String((raw as Any).fillStyle) });
          if (y >= LEGEND_Y_MIN && pendingKey) { legendKeys.push(pendingKey); pendingKey = null; }
          return (raw.fillText as Any)(text, x, y);
        };
      }
      if (p === 'beginPath') return () => { verts = 0; (raw.beginPath as Any)(); };
      if (p === 'moveTo' || p === 'lineTo' || p === 'bezierCurveTo') {
        return (x: number, y: number, ...rest: number[]) => {
          verts += 1; lastY = y;
          return (raw as Any)[p](x, y, ...rest);
        };
      }
      if (p === 'arc') {
        return (cx: number, cy: number, ...rest: number[]) => {
          verts += 1; lastY = cy;
          return (raw.arc as Any)(cx, cy, ...rest);
        };
      }
      if (p === 'fill') return () => { if (inBand()) pendingKey = 'marker'; return (raw.fill as Any)(); };
      if (p === 'fillRect') {
        return (x: number, y: number, w: number, h: number) => {
          if (y >= LEGEND_Y_MIN) pendingKey = 'marker';
          return (raw.fillRect as Any)(x, y, w, h);
        };
      }
      if (p === 'stroke') {
        return () => {
          // A genuine data connecting line spans ≥3 vertices (≥3 data points);
          // a 2-vertex series-colored segment is the legend line-key swatch,
          // not a plot connecting line, so require ≥3 to isolate the plot line.
          if (verts >= 3 && String((raw as Any).strokeStyle).toLowerCase() === matchColor.toLowerCase()) {
            seriesStrokes += 1;
          }
          // A 2-vertex stroke in the legend band is the line-key swatch.
          if (inBand() && verts === 2) pendingKey = 'line';
          return (raw.stroke as Any)();
        };
      }
      const v = (t as Any)[p];
      return typeof v === 'function' ? v.bind(t) : v;
    },
    set(t, p: string, v) { (t as Any)[p] = v; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderChart(proxy, chart, { x: 0, y: 0, w: 640, h: RECT_H }, 1.05);
  return { texts, seriesStrokes, legendKeys };
}

describe.skipIf(!pptxMod || !coreMod || !existsSync(SAMPLE_14))('sample-14 slide-7 pie: white percent-only labels', () => {
  it('renders bare percentages in white (no black category names)', async (context) => {
    const { materializePptxPresentation } = pptxMod as Any;
    const { renderChart } = coreMod as Any;
    const pres = await materializePptxPresentation(readFileSync(SAMPLE_14));
    const slide7 = pres.slides[6];
    const charts: ChartModel[] = [];
    collectCharts(slide7, charts);
    const pie = charts.find(c => c.chartType === 'pie');
    // Private sample names are local conventions and may be reused for a
    // different deck. Treat a structurally different fixture like an absent
    // fixture instead of making the repository-wide test command fail.
    if (!pie) {
      context.skip();
      return;
    }

    const { texts } = renderCapture(pie, renderChart, '#000000');
    // The slice DATA LABELS are the ones drawn in the per-point white; the
    // category names appear only in the LEGEND (drawn in the neutral legend
    // grey, not white). Isolate the white labels — those are the slice labels.
    const whiteLabels = texts.filter(t => t.fill.toLowerCase() === '#ffffff');
    expect(whiteLabels.length, 'pie draws white slice labels').toBeGreaterThan(0);
    for (const l of whiteLabels) {
      // Every white slice label is a bare percentage — no category name, no
      // black text (the pre-fix regression drew black "category + percent").
      expect(l.text).toMatch(/^\d+%$/);
    }
    // No slice label is drawn in black (the overridden series-default color).
    const catNames = pie.categories ?? [];
    const blackSliceLabels = texts.filter(t =>
      t.fill.toLowerCase() === '#000000' && catNames.some(cn => cn && t.text.includes(cn)));
    expect(blackSliceLabels.length, 'no black category-name slice labels').toBe(0);
  });
});

describe.skipIf(!xlsxMod || !coreMod)('sample-30 sheet-1 scatter: markers only (no lines)', () => {
  it('draws no series connecting line for the noFill scatter series', async () => {
    const { materializeXlsxWorksheet } = xlsxMod as Any;
    const { renderChart } = coreMod as Any;
    // Sheet 1 (index 0). parseSheet returns a Worksheet with `charts`.
    const ws = await materializeXlsxWorksheet(readFileSync(SAMPLE_30), 0);
    const anchors = (ws.charts ?? []) as Any[];
    const scatters = anchors.map(a => a.chart as ChartModel).filter(c => c.chartType === 'scatter');
    expect(scatters.length, 'sheet 1 has scatter charts').toBeGreaterThan(0);
    for (const sc of scatters) {
      const color = sc.series[0]?.color ? `#${sc.series[0].color}` : '#4f81bd';
      const { seriesStrokes } = renderCapture(sc, renderChart, color);
      expect(seriesStrokes, 'noFill scatter draws zero connecting lines').toBe(0);
    }
  });

  it('draws a MARKER legend key (not a line swatch) for the markers-only scatter (#803)', async () => {
    const { materializeXlsxWorksheet } = xlsxMod as Any;
    const { renderChart } = coreMod as Any;
    const ws = await materializeXlsxWorksheet(readFileSync(SAMPLE_30), 0);
    const anchors = (ws.charts ?? []) as Any[];
    const scatters = anchors.map(a => a.chart as ChartModel).filter(c => c.chartType === 'scatter');
    expect(scatters.length, 'sheet 1 has scatter charts').toBeGreaterThan(0);
    // The parser must surface the `<a:noFill/>` line override as lineHidden —
    // this is the exact input the legend-key fix keys on.
    expect(
      scatters.some(sc => sc.series.some(s => s.lineHidden === true)),
      'a scatter series is lineHidden (noFill line)',
    ).toBe(true);
    let sawMarker = false;
    for (const sc of scatters) {
      // Force the legend to the bottom so its key lands in the recorded band,
      // independent of the file's own legend visibility (a separate concern).
      const model: ChartModel = { ...sc, showLegend: true, legendPos: 'b' };
      const color = sc.series[0]?.color ? `#${sc.series[0].color}` : '#4f81bd';
      const { legendKeys } = renderCapture(model, renderChart, color);
      // A markers-only scatter must never draw a line-swatch key…
      expect(legendKeys, 'no line-swatch key for a markers-only scatter').not.toContain('line');
      if (legendKeys.includes('marker')) sawMarker = true;
    }
    // …and at least one scatter shows a marker key.
    expect(sawMarker, 'a scatter draws a marker legend key').toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// #931 — §21.2.2.227 varyColors on a single-series bar/column chart. sample-17
// and sample-18 slide 4 each carry a lone column series with NO <c:varyColors>
// element; Office varies each bar by the theme/palette sequence and lists one
// legend entry per data point (four vs. our former one). This probe renders the
// REAL decks and asserts four distinct bar fills plus a per-point legend.
// ─────────────────────────────────────────────────────────────────────────────

/** Render a chart and collect every fillRect color and fillText string. */
function captureBars(
  chart: ChartModel,
  renderChart: (ctx: unknown, c: ChartModel, r: unknown, p?: number) => void,
): { rectFills: string[]; texts: string[] } {
  const canvas = new Canvas(700, 460);
  const raw = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const rectFills: string[] = [];
  const texts: string[] = [];
  const proxy = new Proxy(raw, {
    get(t, p: string) {
      if (p === 'fillRect') {
        return (x: number, y: number, w: number, h: number) => {
          rectFills.push(String((raw as Any).fillStyle).toLowerCase());
          return (raw.fillRect as Any)(x, y, w, h);
        };
      }
      if (p === 'fillText') {
        return (s: string, x: number, y: number) => { texts.push(s); return (raw.fillText as Any)(s, x, y); };
      }
      const v = (t as Any)[p];
      return typeof v === 'function' ? v.bind(t) : v;
    },
    set(t, p: string, v) { (t as Any)[p] = v; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderChart(proxy, chart, { x: 0, y: 0, w: 700, h: 460 }, 1.05);
  return { rectFills, texts };
}

describe.skipIf(!pptxMod || !coreMod || varySamples.length === 0)(
  '#931 single-series bar varyColors (sample-17/18 slide 4)',
  () => {
    it('renders four distinct bar colors and a per-point legend', async () => {
      const { materializePptxPresentation } = pptxMod as Any;
      const { renderChart } = coreMod as Any;
      for (const path of varySamples) {
        const pres = await materializePptxPresentation(readFileSync(path));
        const slide4 = pres.slides[3];
        const charts: ChartModel[] = [];
        collectCharts(slide4, charts);
        const bar = charts.find(
          (c) => typeof c.chartType === 'string' && /Bar/.test(c.chartType) && c.series.length === 1,
        );
        expect(bar, `${path}: slide 4 has a single-series bar`).toBeTruthy();
        const chart = bar as ChartModel;
        // The parser flags default vary-by-point (element absent) as varyColors.
        expect(chart.varyColors, 'single-series bar varies by default').toBe(true);

        const { rectFills, texts } = captureBars(chart, renderChart);
        const distinct = [...new Set(rectFills)];
        // Four categories → four distinct bar fills (was one before the fix).
        expect(
          distinct.length,
          `${path}: at least ${chart.categories.length} distinct bar colors`,
        ).toBeGreaterThanOrEqual(chart.categories.length);
        // Each category label now appears twice — once on the axis, once in the
        // per-point legend (the pre-fix legend held only the series name).
        for (const cat of chart.categories) {
          if (!cat) continue;
          const occurrences = texts.filter((s) => s === cat).length;
          expect(occurrences, `${path}: "${cat}" drawn as axis label + legend entry`).toBeGreaterThanOrEqual(2);
        }
      }
    });
  },
);
