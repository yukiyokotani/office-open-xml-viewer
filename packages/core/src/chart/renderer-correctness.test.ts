// Behavioral tests for the chart-correctness fixes:
//   CH1 — bar/column negative values extend downward from the zero line
//   CH2 — stackedLine / stackedLinePct series are cumulatively stacked
//   CH3 — tick / data labels use the locale-independent §18.8.30 formatter
//
// These assert observable geometry (fillRect bounds, gridline label text)
// captured through a lightweight recording context, complementing the
// draw-call-signature characterization test.

import { describe, it, expect, vi } from 'vitest';
import type {
  ChartModel,
  ChartRect,
  ChartSeries,
  ChartThreeDPictureOptions,
} from '../types/chart';
import {
  chartLabelPaintWorkCount,
  classicMarkerPaintWorkCount,
  renderChart as renderChartCore,
} from './renderer.js';
import {
  chartExHierarchyLabelPaintWorkCount,
  renderChartExChart,
} from './chart-ex-renderer.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';
import { formatChartValWithCode } from './chart-number-format.js';
import { BOX_WHISKER_SLOT_GUTTER_FRACTION } from './box-whisker.js';
import {
  chartImageFillKey,
  chartImageFillPaintWork,
  collectChartMarkerImageFills,
  collectChartMarkerImageFillsForCharts,
} from './image-fill.js';
import { MAX_CANVAS_CHART_POINTS, sourceChartStructureCount } from './resource-limits.js';

const testThreeD = { render: renderSimpleThreeDChart };
const testChartEx = { render: renderChartExChart };

it('uses collision-free tuple keys for decoded chart image sources', () => {
  expect(chartImageFillKey({
    fillType: 'image', stretch: true, imagePath: 'a|b', mimeType: 'image/png',
  })).not.toBe(chartImageFillKey({
    fillType: 'image', stretch: true, imagePath: 'a', svgImagePath: 'b',
    mimeType: 'image/png',
  }));
});

it('prefetches and paints direct chart-space and plot-area picture fills', () => {
  const chartPicture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/chart-background.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const plotPicture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/plot-background.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const model = baseModel({
    chartFill: chartPicture,
    chartFillPaintAuthored: true,
    plotAreaFill: plotPicture,
    plotAreaFillPaintAuthored: true,
    series: [series({ values: [1] })],
  });
  const bitmap = { width: 8, height: 8 } as unknown as CanvasImageSource;

  expect(collectChartMarkerImageFills(model)).toEqual([chartPicture, plotPicture]);
  const rec = recordingCtx();
  renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
  expect(rec.drawImages).toHaveLength(2);
  expect(rec.drawImages.every(call => call[0] === bitmap)).toBe(true);
});

it('collects only the effective direct bubble point picture', () => {
  const seriesPicture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/series.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const pointPicture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/point.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const fills = collectChartMarkerImageFills(baseModel({
    chartType: 'bubble',
    categories: ['0', '1'],
    series: [series({
      values: [0.25, 0.75],
      bubbleSizes: [100, 100],
      chartexStyle: { fillPaints: [seriesPicture], fillPaintAuthored: true },
      dataPointOverrides: [
        { idx: 0, chartexStyle: { fillPaints: [pointPicture], fillPaintAuthored: true } },
        { idx: 1, chartexStyle: { fillHidden: true, fillPaintAuthored: true } },
      ],
    })],
  }));
  expect(fills.map(fill => fill.imagePath)).toEqual(['xl/media/point.png']);
});

it('collects the linked bubble picture selected by point index after direct colors', () => {
  const linkedPictures = [0, 1].map(index => ({
    fillType: 'image' as const,
    imagePath: `xl/media/linked-${index}.png`,
    mimeType: 'image/png',
    stretch: true,
  }));
  const fills = collectChartMarkerImageFills(baseModel({
    chartType: 'bubble',
    categories: ['0', '1'],
    chartStyleRoles: {
      dataPoint: { fillPaints: linkedPictures, fillPaintAuthored: true },
    },
    series: [series({
      values: [0.25, 0.75],
      bubbleSizes: [100, 100],
      dataPointColors: ['FF0000', null],
    })],
  }));
  expect(fills.map(fill => fill.imagePath)).toEqual(['xl/media/linked-1.png']);
});

it('does not prefetch bubble pictures for points whose sizes cannot paint', () => {
  const picture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/invisible.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const model = baseModel({
    chartType: 'bubble',
    categories: ['0', '1', '2'],
    chartStyleRoles: {
      dataPoint: { fillPaints: [picture], fillPaintAuthored: true },
    },
    series: [series({ values: [1, 2, 3], bubbleSizes: [0, null, -1] })],
  });
  expect(collectChartMarkerImageFills(model)).toEqual([]);
  expect(collectChartMarkerImageFills({ ...model, bubbleScale: 0 })).toEqual([]);
  expect(collectChartMarkerImageFills({ ...model, showNegativeBubbles: true })).toEqual([]);
});

it('uses the owning bubble group settings for prefetch and paint work', () => {
  const picture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/group-bubble.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const model = baseModel({
    chartType: 'scatter',
    categories: ['1'],
    chartStyleRoles: {
      dataPoint: { fillPaints: [picture], fillPaintAuthored: true },
    },
    series: [
      series({ values: [1], seriesType: 'scatter', markerSymbol: 'none' }),
      series({ values: [2], seriesType: 'scatter', bubbleSizes: [25] }),
    ],
    plotGroups: [
      plotGroup('scatter', 0, 1, { scatterStyle: 'line' }),
      plotGroup('bubble', 1, 1, { bubbleScale: 0, showNegativeBubbles: true }),
    ],
  });
  const bitmap = { width: 8, height: 8 } as unknown as CanvasImageSource;
  expect(collectChartMarkerImageFills(model)).toEqual([]);
  expect(classicMarkerPaintWorkCount(model, () => bitmap, 1, RECT)).toBe(0);

  const visible = {
    ...model,
    plotGroups: [
      plotGroup('scatter', 0, 1, { scatterStyle: 'line' }),
      plotGroup('bubble', 1, 1, { bubbleScale: 100, showNegativeBubbles: true }),
    ],
  };
  expect(collectChartMarkerImageFills(visible)).toEqual([picture]);
  expect(classicMarkerPaintWorkCount(visible, () => bitmap, 1, RECT)).toBe(1);
  const rec = recordingCtx();
  renderChartCore(rec.ctx, visible, RECT, 1, 0, testThreeD, undefined, () => bitmap);
  expect(rec.drawImages).toHaveLength(1);
});

it('prefetches and paints one bubble picture for both plot and 3-D legend key', () => {
  const picture = {
    fillType: 'image' as const,
    imagePath: 'xl/media/bubble-key.png',
    mimeType: 'image/png',
    stretch: true,
  };
  const model = baseModel({
    chartType: 'bubble', showLegend: true, categories: ['0'],
    series: [series({
      values: [1], bubbleSizes: [100], bubble3D: true,
      chartexStyle: {
        fillPaints: [picture], fillPaintAuthored: true,
      },
    })],
    catAxisMin: 0, catAxisMax: 1, valMin: 0, valMax: 2,
  });
  const bitmap = { width: 8, height: 8 } as unknown as CanvasImageSource;
  expect(collectChartMarkerImageFills(model)).toEqual([picture]);
  expect(classicMarkerPaintWorkCount(model, () => bitmap, 1, RECT)).toBe(32);
  const rec = recordingCtx();
  renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
  expect(rec.drawImages).toHaveLength(2);
  expect(rec.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(6);
});
const renderChart: typeof renderChartCore = (
  ctx,
  chart,
  rect,
  ptToPx,
  shapeRotationDeg,
  threeD = testThreeD,
  regionMap,
  imageLookup,
  chartEx = testChartEx,
) => renderChartCore(
  ctx,
  chart,
  rect,
  ptToPx,
  shapeRotationDeg,
  threeD,
  regionMap,
  imageLookup,
  chartEx,
);

interface RectCall { x: number; y: number; w: number; h: number; fs: string }
interface StrokeRectCall {
  x: number; y: number; w: number; h: number; ss: string; lw: number;
  dash: number[]; cap: string; join: string;
}
interface TextCall {
  text: string;
  x: number;
  y: number;
  align: string;
  baseline: string;
  font?: string;
  width?: number;
  fillStyle?: string;
}

interface Recorded {
  ctx: CanvasRenderingContext2D;
  rects: RectCall[];
  strokeRects: StrokeRectCall[];
  texts: TextCall[];
  clips: Array<{ x: number; y: number; w: number; h: number }>;
  clipCalls: number;
  quadratics: Array<{ cpx: number; cpy: number; x: number; y: number }>;
  gradients: Array<{
    kind: 'linear' | 'radial';
    args: number[];
    stops: Array<{ position: number; color: string }>;
  }>;
  compositeModes: string[];
  arcs: Array<{ x: number; y: number; r: number }>;
  ellipses: Array<{ x: number; y: number; rx: number; ry: number }>;
  rotations: number[];
  translations: Array<{ x: number; y: number }>;
  drawImages: unknown[][];
  filledPaths: Array<{ points: Array<{ x: number; y: number }>; fillStyle: string }>;
  strokedPaths: Array<{ points: Array<{ x: number; y: number }>; strokeStyle: string }>;
  strokeDetails: Array<{
    strokeStyle: string; lineWidth: number; dash: number[]; cap: string; join: string;
  }>;
  paintEvents: Array<
    | { kind: 'stroke'; strokeStyle: string }
    | { kind: 'fill'; fillStyle: string }
    | { kind: 'rect'; fillStyle: string }
    | { kind: 'text'; text: string }
  >;
}

type FillPaintEvent = Extract<Recorded['paintEvents'][number], { kind: 'fill' }>;

/** Mesh materials scale all channels of the authored base color by one normal-
 * light factor. Tests identify that semantic color family without freezing a
 * particular camera/material coefficient. */
function isMaterialColor(fillStyle: string, baseHex: string): boolean {
  const actual = /^#([0-9a-f]{6})$/i.exec(fillStyle)?.[1];
  const base = /^#?([0-9a-f]{6})$/i.exec(baseHex)?.[1];
  if (!actual || !base) return false;
  const actualChannels = [0, 2, 4].map(offset => parseInt(actual.slice(offset, offset + 2), 16));
  const baseChannels = [0, 2, 4].map(offset => parseInt(base.slice(offset, offset + 2), 16));
  const factors: number[] = [];
  for (let index = 0; index < 3; index++) {
    if (baseChannels[index] === 0) {
      if (actualChannels[index] !== 0) return false;
    } else {
      factors.push(actualChannels[index] / baseChannels[index]);
    }
  }
  if (!factors.length || factors.some(factor => factor < 0.55 || factor > 1.01)) return false;
  return Math.max(...factors) - Math.min(...factors) < 0.04;
}

function materialFills(rec: Recorded, baseHex: string): FillPaintEvent[] {
  return rec.paintEvents.filter((event): event is FillPaintEvent =>
    event.kind === 'fill' && isMaterialColor(event.fillStyle, baseHex));
}

function isSurfaceMaterialColor(fillStyle: string, baseHex: string): boolean {
  const actual = /^#([0-9a-f]{6})$/i.exec(fillStyle)?.[1];
  const base = /^#?([0-9a-f]{6})$/i.exec(baseHex)?.[1];
  if (!actual || !base) return false;
  const factors: number[] = [];
  for (const offset of [0, 2, 4]) {
    const source = parseInt(base.slice(offset, offset + 2), 16);
    const painted = parseInt(actual.slice(offset, offset + 2), 16);
    if (source === 0) {
      if (painted !== 0) return false;
    } else if (painted < 255) {
      factors.push(painted / source);
    }
  }
  return factors.length > 0
    && factors.every(factor => factor >= 0.45 && factor <= 1.25)
    && Math.max(...factors) - Math.min(...factors) < 0.05;
}

/** Minimal recording 2D context: captures fillRect + fillText, tracks the
 *  handful of state props the renderer reads, and models text width. */
function recordingCtx(measureOverride?: (text: string, fontPx: number) => number | null): Recorded {
  const rects: RectCall[] = [];
  const strokeRects: StrokeRectCall[] = [];
  const texts: TextCall[] = [];
  const clips: Array<{ x: number; y: number; w: number; h: number }> = [];
  let clipCalls = 0;
  const quadratics: Recorded['quadratics'] = [];
  const gradients: Recorded['gradients'] = [];
  const compositeModes: string[] = [];
  const arcs: Recorded['arcs'] = [];
  const ellipses: Recorded['ellipses'] = [];
  const rotations: number[] = [];
  const translations: Recorded['translations'] = [];
  const drawImages: unknown[][] = [];
  const filledPaths: Recorded['filledPaths'] = [];
  const strokedPaths: Recorded['strokedPaths'] = [];
  const strokeDetails: Recorded['strokeDetails'] = [];
  const paintEvents: Recorded['paintEvents'] = [];
  let dash: number[] = [];
  let pathRect: { x: number; y: number; w: number; h: number } | null = null;
  let pathPoints: Array<{ x: number; y: number }> = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
    globalCompositeOperation: 'source-over',
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const textWidth = (text: string): number => {
    const px = fontPx(String(state.font));
    const overridden = measureOverride?.(String(text), px);
    if (overridden != null) return overridden;
    let w = 0;
    for (const ch of String(text)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
    return w;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => ({ width: textWidth(t) });
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) => {
            paintEvents.push({ kind: 'rect', fillStyle: String(state.fillStyle) });
            rects.push({ x, y, w, h, fs: String(state.fillStyle) });
          };
        case 'fillText':
          return (text: string, x: number, y: number) => {
            paintEvents.push({ kind: 'text', text: String(text) });
            texts.push({
              text,
              x,
              y,
              align: String(state.textAlign),
              baseline: String(state.textBaseline),
              font: String(state.font),
              width: textWidth(text),
              fillStyle: String(state.fillStyle),
            });
          };
        case 'strokeRect':
          return (x: number, y: number, w: number, h: number) =>
            strokeRects.push({
              x, y, w, h, ss: String(state.strokeStyle), lw: Number(state.lineWidth),
              dash: [...dash], cap: String(state.lineCap), join: String(state.lineJoin),
            });
        case 'drawImage':
          return (...args: unknown[]) => { drawImages.push(args); };
        case 'createLinearGradient':
        case 'createRadialGradient':
          return (...args: number[]) => {
            const gradient = {
              kind: prop === 'createRadialGradient' ? 'radial' as const : 'linear' as const,
              args,
              stops: [] as Array<{ position: number; color: string }>,
            };
            gradients.push(gradient);
            return {
              addColorStop(position: number, color: string) {
                gradient.stops.push({ position, color });
              },
            };
          };
        case 'beginPath':
          return () => { pathRect = null; pathPoints = []; };
        case 'rect':
          return (x: number, y: number, w: number, h: number) => { pathRect = { x, y, w, h }; };
        case 'clip':
          return () => { clipCalls++; if (pathRect) clips.push(pathRect); };
        case 'arc':
          return (x: number, y: number, r: number) => { arcs.push({ x, y, r }); };
        case 'ellipse':
          return (x: number, y: number, rx: number, ry: number) => {
            ellipses.push({ x, y, rx, ry });
          };
        case 'save': case 'restore': case 'closePath':
          return () => paintEvents.push({
            kind: 'stroke', strokeStyle: String(state.strokeStyle),
          });
        case 'stroke':
          return () => {
            const strokeStyle = String(state.strokeStyle);
            paintEvents.push({ kind: 'stroke', strokeStyle });
            if (pathPoints.length > 0) {
              strokedPaths.push({
                points: pathPoints.map(point => ({ ...point })),
                strokeStyle,
              });
            }
            strokeDetails.push({
              strokeStyle,
              lineWidth: Number(state.lineWidth),
              dash: [...dash],
              cap: String(state.lineCap),
              join: String(state.lineJoin),
            });
          };
        case 'fill':
          return () => {
            paintEvents.push({ kind: 'fill', fillStyle: String(state.fillStyle) });
            if (pathPoints.length >= 3) {
              filledPaths.push({
                points: pathPoints.map(point => ({ ...point })),
                fillStyle: String(state.fillStyle),
              });
            }
          };
        case 'moveTo': case 'lineTo':
          return (x: number, y: number) => { pathPoints.push({ x, y }); };
        case 'quadraticCurveTo':
          return (cpx: number, cpy: number, x: number, y: number) => {
            quadratics.push({ cpx, cpy, x, y });
          };
        case 'bezierCurveTo':
          return () => undefined;
        case 'setLineDash':
          return (value: number[] = []) => { dash = [...value]; };
        case 'getLineDash':
          return () => [...dash];
        case 'translate':
          return (x: number, y: number) => { translations.push({ x, y }); };
        case 'rotate':
          return (angle: number) => { rotations.push(angle); };
        case 'clearRect': case 'strokeText':
        case 'scale':
        case 'setTransform': case 'resetTransform':
          return () => undefined;
        case 'getTransform':
          return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) {
      state[prop] = value;
      if (prop === 'globalCompositeOperation') compositeModes.push(String(value));
      return true;
    },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    rects,
    strokeRects,
    texts,
    clips,
    get clipCalls() { return clipCalls; },
    quadratics,
    gradients,
    compositeModes,
    arcs,
    ellipses,
    rotations,
    translations,
    drawImages,
    filledPaths,
    strokedPaths,
    strokeDetails,
    paintEvents,
  };
}

function baseModel(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: [],
    series: [],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: false,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out',
    catAxisMajorTickMark: 'out',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: null,
    valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    ...over,
  };
}

function series(over: Partial<ChartSeries>): ChartSeries {
  return { name: '', color: null, values: [], ...over };
}

function plotGroup(
  kind: NonNullable<ChartModel['plotGroups']>[number]['kind'],
  seriesStart: number,
  seriesCount: number,
  over: Partial<NonNullable<ChartModel['plotGroups']>[number]> = {},
): NonNullable<ChartModel['plotGroups']>[number] {
  return {
    kind,
    seriesStart,
    seriesCount,
    categoryAxis: 'primary',
    valueAxis: 'primary',
    seriesAxis: 'none',
    ...over,
  };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

it('keeps classic 2-D charts in the default renderer and ChartEx opt-in', () => {
  const classic = recordingCtx();
  renderChartCore(classic.ctx, baseModel({
    chartType: 'line',
    categories: ['A', 'B'],
    series: [series({ values: [1, 2] })],
  }), RECT, 1);
  expect(classic.texts.map(item => item.text)).not.toContain('Unsupported chart');

  const omitted = recordingCtx();
  const waterfall = baseModel({
    chartType: 'waterfall',
    categories: ['A', 'B'],
    series: [series({ values: [1, 2] })],
  });
  renderChartCore(omitted.ctx, waterfall, RECT, 1);
  expect(omitted.texts.map(item => item.text)).toContain('Unsupported chart');

  const enabled = recordingCtx();
  renderChartCore(
    enabled.ctx,
    waterfall,
    RECT,
    1,
    0,
    undefined,
    undefined,
    undefined,
    testChartEx,
  );
  expect(enabled.texts.map(item => item.text)).not.toContain('Unsupported chart');
  expect(enabled.rects.length).toBeGreaterThan(0);
});

describe('ordered classic plot groups', () => {
  it.each([
    {
      kind: 'line' as const,
      chartType: 'line',
      group: { grouping: 'standard' },
      seriesType: 'line',
    },
    {
      kind: 'area' as const,
      chartType: 'area',
      group: { grouping: 'standard' },
      seriesType: 'area',
    },
    {
      kind: 'bar' as const,
      chartType: 'clusteredBar',
      group: { grouping: 'clustered', barDirection: 'col' },
      seriesType: 'bar',
    },
    {
      kind: 'scatter' as const,
      chartType: 'scatter',
      group: { scatterStyle: 'lineMarker' },
      seriesType: 'scatter',
    },
  ])('keeps a single $kind group byte-equivalent to the legacy projection', entry => {
    const base = baseModel({
      chartType: entry.chartType,
      scatterStyle: entry.kind === 'scatter' ? 'lineMarker' : null,
      categories: entry.kind === 'scatter' ? ['1', '2'] : ['A', 'B'],
      series: [series({
        values: [1, 2], seriesType: entry.seriesType,
        showMarker: true, markerSymbol: 'circle',
      })],
    });
    const signature = (model: ChartModel) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, model, RECT, 1);
      return {
        rects: rec.rects,
        strokes: rec.strokeDetails,
        paths: rec.strokedPaths,
        fills: rec.filledPaths,
        arcs: rec.arcs,
        texts: rec.texts,
      };
    };
    expect(signature({
      ...base,
      plotGroups: [plotGroup(entry.kind, 0, 1, entry.group)],
    })).toEqual(signature(base));
  });

  it('keeps a single negative column group byte-equivalent to the legacy projection', () => {
    const base = baseModel({
      chartType: 'clusteredBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr'],
      series: [series({
        values: [150, -300, 450, -120],
        seriesType: 'bar',
        invertIfNegative: true,
      })],
    });
    const signature = (model: ChartModel) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, model, RECT, 1);
      return {
        rects: rec.rects,
        strokes: rec.strokeDetails,
        paths: rec.strokedPaths,
        fills: rec.filledPaths,
        texts: rec.texts,
      };
    };
    expect(signature({
      ...base,
      series: base.series.map(entry => ({
        ...entry,
        barGroupIndex: 0,
        barGroupGrouping: 'clustered',
        barGroupDirection: 'col',
      })),
      plotGroups: [plotGroup('bar', 0, 1, {
        grouping: 'clustered', barDirection: 'col',
      })],
    })).toEqual(signature(base));
  });

  it('fails closed before painting a mixed 2-D and 3-D plot', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [
        series({ color: 'FF0000', values: [2], seriesType: 'bar' }),
        series({ color: '0000FF', values: [3], seriesType: 'bar' }),
      ],
      plotGroups: [plotGroup('bar', 0, 1), plotGroup('bar3D', 1, 1)],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('Unsupported chart');
    expect(rec.rects).toEqual([]);
    expect(rec.filledPaths).toEqual([]);
  });

  it('composites scatter and bubble groups in source order', () => {
    const render = (bubbleFirst: boolean): Recorded => {
      const scatter = series({
        color: '0066CC', values: [1, 3], categories: ['0', '2'],
        seriesType: 'scatter', showMarker: false,
      });
      const bubble = series({
        color: 'FF8800', values: [2], categories: ['1'], bubbleSizes: [100],
        seriesType: 'scatter', markerSymbol: 'circle',
      });
      const ordered = bubbleFirst ? [bubble, scatter] : [scatter, bubble];
      const groups = bubbleFirst
        ? [
            plotGroup('bubble', 0, 1, { categoryAxis: 'secondary', valueAxis: 'secondary' }),
            plotGroup('scatter', 1, 1, { scatterStyle: 'line' }),
          ]
        : [
            plotGroup('scatter', 0, 1, { scatterStyle: 'line' }),
            plotGroup('bubble', 1, 1, { categoryAxis: 'secondary', valueAxis: 'secondary' }),
          ];
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'scatter', categories: ['0', '2'], series: ordered, plotGroups: groups,
        secondaryCatAxis: {
          min: 0, max: 2, title: null, hidden: false, lineHidden: true,
          majorTickMark: 'none',
        },
        secondaryValAxis: {
          min: 0, max: 4, title: null, hidden: false, lineHidden: true,
          majorTickMark: 'none',
        },
      }), RECT, 1);
      return rec;
    };
    const eventOrder = (rec: Recorded): [number, number] => [
      rec.paintEvents.findIndex(event => event.kind === 'stroke' && event.strokeStyle === '#0066CC'),
      rec.paintEvents.findIndex(event => event.kind === 'fill' && event.fillStyle === '#FF8800'),
    ];
    const scatterFirst = eventOrder(render(false));
    const bubbleFirst = eventOrder(render(true));
    expect(scatterFirst[0]).toBeGreaterThanOrEqual(0);
    expect(scatterFirst[1]).toBeGreaterThan(scatterFirst[0]);
    expect(bubbleFirst[1]).toBeGreaterThanOrEqual(0);
    expect(bubbleFirst[0]).toBeGreaterThan(bubbleFirst[1]);
  });

  it('isolates stacked values between separate line groups', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A'],
      valMin: 0,
      valMax: 30,
      series: [
        series({ values: [10], markerSymbol: 'circle', seriesType: 'line' }),
        series({ values: [10], markerSymbol: 'circle', seriesType: 'line' }),
        series({ values: [2], markerSymbol: 'circle', seriesType: 'line' }),
      ],
      plotGroups: [
        plotGroup('line', 0, 2, { grouping: 'stacked' }),
        plotGroup('line', 2, 1, { grouping: 'standard' }),
      ],
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(3);
    expect(rec.arcs[1].y).toBeLessThan(rec.arcs[0].y);
    expect(rec.arcs[2].y).toBeGreaterThan(rec.arcs[0].y);
    expect(rec.arcs[2].y).toBeGreaterThan(rec.arcs[1].y);
  });

  it.each(['line', 'area'] as const)(
    'keeps a secondary percent-stacked %s group on its own normalized axis',
    kind => {
      const renderSecondary = (values: [number, number], grouping: string): Recorded => {
        const rec = recordingCtx();
        renderChart(rec.ctx, baseModel({
          chartType: kind,
          categories: ['A'],
          valMin: 0,
          valMax: 100,
          secondaryValAxis: {
            min: 0, max: 1, title: null, hidden: false, lineHidden: false,
            majorTickMark: 'none', formatCode: '0%',
          },
          series: [
            series({ values: [100], markerSymbol: 'none', seriesType: kind }),
            series({ values: [values[0]], markerSymbol: 'circle', seriesType: kind }),
            series({ values: [values[1]], markerSymbol: 'circle', seriesType: kind }),
          ],
          plotGroups: [
            plotGroup(kind, 0, 1, { grouping: 'standard' }),
            plotGroup(kind, 1, 2, { grouping, valueAxis: 'secondary' }),
          ],
        }), RECT, 1);
        return rec;
      };
      const percent = renderSecondary([60, 40], 'percentStacked');
      const normalized = renderSecondary([0.6, 1], 'standard');
      expect(percent.arcs).toHaveLength(2);
      expect(percent.arcs.map(arc => arc.y)).toEqual(normalized.arcs.map(arc => arc.y));
    },
  );

  it('keeps a percent-stacked area group in ratio space on a raw shared axis', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      valMin: 0,
      valMax: 100,
      series: [
        series({ color: '666666', values: [100], seriesType: 'bar', barGroupIndex: 0 }),
        series({ color: 'FF0000', values: [60], seriesType: 'area', showMarker: true }),
        series({ color: '00AA00', values: [40], seriesType: 'area', showMarker: true }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { grouping: 'clustered', barDirection: 'col' }),
        plotGroup('area', 1, 2, { grouping: 'percentStacked' }),
      ],
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(2);
    expect(rec.arcs.every(arc => arc.y > RECT.h * 0.65)).toBe(true);
    expect(rec.arcs[1].y).toBeLessThan(rec.arcs[0].y);
  });

  it.each([
    { chartType: 'clusteredBar', firstDirection: 'col', secondDirection: 'bar' },
    { chartType: 'clusteredBarH', firstDirection: 'bar', secondDirection: 'col' },
  ] as const)('retains mixed bar directions over the first group axes ($chartType)', ({
    chartType, firstDirection, secondDirection,
  }) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['A', 'B', 'C', 'D', 'E', 'F'],
      valMin: 0,
      valMax: 5,
      series: [
        series({ color: 'FF0000', values: [4, 4, 4, 4, 4, 4], seriesType: 'bar', barGroupIndex: 0, barGroupDirection: firstDirection }),
        series({ color: '0000FF', values: [3, 3, 3, 3, 3, 3], seriesType: 'bar', barGroupIndex: 1, barGroupDirection: secondDirection }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { grouping: 'clustered', barDirection: firstDirection }),
        plotGroup('bar', 1, 1, { grouping: 'clustered', barDirection: secondDirection }),
      ],
    }), RECT, 1);
    const red = rec.rects.find(rect => rect.fs === '#FF0000');
    const blue = rec.rects.find(rect => rect.fs === '#0000FF');
    expect(red).toBeDefined();
    expect(blue).toBeDefined();
    const first = red as NonNullable<typeof red>;
    const second = blue as NonNullable<typeof blue>;
    expect(firstDirection === 'bar' ? first.w > first.h : first.h > first.w).toBe(true);
    expect(secondDirection === 'bar' ? second.w > second.h : second.h > second.w).toBe(true);
    expect(rec.texts.map(item => item.text)).not.toContain('Unsupported chart');
  });

  it('does not extrapolate mixed bar-direction ownership beyond the observed pair', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [
        series({ values: [1], seriesType: 'bar' }),
        series({ values: [2], seriesType: 'bar' }),
        series({ values: [3], seriesType: 'bar' }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { barDirection: 'col' }),
        plotGroup('bar', 1, 1, { barDirection: 'bar' }),
        plotGroup('bar', 2, 1, { barDirection: 'col' }),
      ],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('Unsupported chart');
    expect(rec.rects).toEqual([]);
  });

  it.each([
    { overlayKind: 'line' as const, overlayEvent: 'stroke' as const },
    { overlayKind: 'area' as const, overlayEvent: 'fill' as const },
  ])('keeps the observed fixed bar/$overlayKind family layer after reversed XML order', ({
    overlayKind, overlayEvent,
  }) => {
    const rec = recordingCtx();
    const overlay = series({
      color: 'FF0000', lineColor: 'FF0000', values: [2, 3],
      seriesType: overlayKind, showMarker: false,
    });
    const bar = series({
      color: '0000FF', values: [1, 2], seriesType: 'bar', barGroupIndex: 1,
    });
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A', 'B'],
      // The overlay precedes the bar in XML/model order. Office nevertheless
      // keeps line above bars and area behind bars for these observed pairs.
      series: [overlay, bar],
      plotGroups: [
        plotGroup(overlayKind, 0, 1, { grouping: 'standard' }),
        plotGroup('bar', 1, 1, { grouping: 'clustered', barDirection: 'col' }),
      ],
    }), RECT, 1);
    const overlayIndex = rec.paintEvents.findIndex(event => overlayEvent === 'fill'
      ? event.kind === 'fill' && event.fillStyle === '#FF0000'
      : event.kind === 'stroke' && event.strokeStyle === '#FF0000');
    const barIndex = rec.paintEvents.findIndex(event =>
      event.kind === 'rect' && event.fillStyle === '#0000FF'
    );
    expect(overlayIndex).toBeGreaterThanOrEqual(0);
    expect(barIndex).toBeGreaterThanOrEqual(0);
    if (overlayKind === 'line') expect(overlayIndex).toBeGreaterThan(barIndex);
    else expect(overlayIndex).toBeLessThan(barIndex);
  });

  it.each(['line', 'area'] as const)(
    'fails closed for an unimplemented secondary category-axis %s group',
    kind => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: kind,
        categories: ['A'],
        secondaryCatAxis: {
          min: null, max: null, title: null, hidden: false, lineHidden: false,
          majorTickMark: 'none',
        },
        secondaryValAxis: {
          min: 0, max: 2, title: null, hidden: false, lineHidden: false,
          majorTickMark: 'none',
        },
        series: [
          series({ values: [1], seriesType: kind }),
          series({ values: [1], seriesType: kind, useSecondaryAxis: true }),
        ],
        plotGroups: [
          plotGroup(kind, 0, 1, { grouping: 'standard' }),
          plotGroup(kind, 1, 1, {
            categoryAxis: 'secondary', valueAxis: 'secondary', grouping: 'standard',
          }),
        ],
      }), RECT, 1);
      expect(rec.texts.map(item => item.text)).toContain('Unsupported chart');
      expect(rec.strokedPaths).toEqual([]);
      expect(rec.rects).toEqual([]);
    },
  );

  it('fails closed when public empty-group metadata disagrees with the visible family', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      series: [series({ values: [1] })],
      plotGroups: [plotGroup('bar', 0, 0), plotGroup('pie', 0, 1)],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('Unsupported chart');
  });

  it('keeps stock role ownership and paints a later line group', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      stockHiLowLines: true,
      stockHiLowLineColor: '123456',
      series: [
        series({ values: [5, 6], seriesType: 'stock' }),
        series({ values: [1, 2], seriesType: 'stock' }),
        series({ values: [3, 4], seriesType: 'stock' }),
        series({
          color: 'FF0000', lineColor: 'FF0000', values: [2, 5],
          seriesType: 'line', showMarker: false,
          chartexStyle: { lineDash: 'dash', lineCap: 'rnd', lineJoin: 'bevel' },
        }),
      ],
      plotGroups: [plotGroup('stock', 0, 3), plotGroup('line', 3, 1)],
    }), RECT, 1);
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#123456' });
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
    const hiLow = rec.strokedPaths.find(path => path.strokeStyle === '#123456'
      && path.points.length === 2 && path.points[0].x === path.points[1].x);
    const overlay = rec.strokedPaths.find(path => path.strokeStyle === '#FF0000'
      && path.points.length >= 2 && path.points[0].x !== path.points.at(-1)?.x);
    expect(hiLow).toBeDefined();
    expect(overlay).toBeDefined();
    expect(rec.strokeDetails.some(detail => detail.strokeStyle === '#FF0000'
      && detail.dash.length > 0 && detail.cap === 'round' && detail.join === 'bevel')).toBe(true);
  });

  it('bounds group metadata and rejects malformed public ranges', () => {
    const emptyGroups = Array.from(
      { length: MAX_CANVAS_CHART_POINTS + 1 },
      () => plotGroup('pie', 0, 0),
    );
    expect(sourceChartStructureCount(baseModel({ plotGroups: emptyGroups })))
      .toBe(MAX_CANVAS_CHART_POINTS + 1);
    expect(sourceChartStructureCount(baseModel({
      series: [series({ values: [1] })],
      plotGroups: [plotGroup('line', 1, 1)],
    }))).toBe(MAX_CANVAS_CHART_POINTS + 1);
  });

  it('plans many line groups without rescanning every group per group', () => {
    let axisReads = 0;
    const groupCount = 200;
    const groups = Array.from({ length: groupCount }, (_, index) => ({
      kind: 'line' as const,
      seriesStart: index,
      seriesCount: 1,
      categoryAxis: 'primary' as const,
      get valueAxis() { axisReads++; return 'primary' as const; },
      seriesAxis: 'none' as const,
      grouping: index % 2 === 0 ? 'standard' : 'stacked',
    }));
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'], plotGroups: groups,
      series: Array.from({ length: groupCount }, (_, index) => series({
        values: [index + 1], seriesType: 'line', lineHidden: true,
        showMarker: false, markerSymbol: 'none',
      })),
    }), RECT, 1);
    expect(axisReads).toBeLessThan(groupCount * 12);
  });
});

describe('chart-space background', () => {
  it('fills the complete chart rectangle, including the axis-label gutters', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      chartBg: 'F2F2F2',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
    }), RECT, 1);

    expect(rec.rects[0]).toEqual({ x: 0, y: 0, w: 640, h: 360, fs: '#F2F2F2' });
  });

  it('clips fill and chart content to one rounded path and strokes the same geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      roundedCorners: true,
      chartBg: 'F2F2F2',
      chartBorderColor: '0055AA',
      chartBorderWidthEmu: 25_400,
    }), RECT, 1);

    expect(rec.clipCalls).toBe(1);
    expect(rec.rects[0]).toEqual({ x: 0, y: 0, w: 640, h: 360, fs: '#F2F2F2' });
    expect(rec.strokeRects).toHaveLength(0);
    // Four corners for the outer clip plus four for the inset border.
    expect(rec.quadratics).toHaveLength(8);
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#0055AA' });
  });

  it('keeps the application-defined rounded radius in points and clamps tiny frames', () => {
    for (const [w, h] of [[390, 315], [570, 240], [285, 420]]) {
      const aspect = recordingCtx();
      renderChart(aspect.ctx, baseModel({
        roundedCorners: true,
        chartBg: 'F2F2F2',
      }), { x: 0, y: 0, w, h }, 1);
      expect(aspect.quadratics[0]?.y).toBe(10);
    }

    const scaled = recordingCtx();
    renderChart(scaled.ctx, baseModel({
      roundedCorners: true,
      chartBg: 'F2F2F2',
    }), { x: 0, y: 0, w: 200, h: 100 }, 2);
    // Desktop Excel vector output fixes the radius at 10pt across aspect ratios.
    expect(scaled.quadratics[0]?.y).toBe(20);

    const tiny = recordingCtx();
    renderChart(tiny.ctx, baseModel({
      roundedCorners: true,
      chartBg: 'F2F2F2',
    }), { x: 0, y: 0, w: 8, h: 6 }, 2);
    // Geometry, not a sample-specific threshold, bounds the radius to h/2.
    expect(tiny.quadratics[0]?.y).toBe(3);
  });

  it('keeps both compound rails inside the rounded chart-space clip', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      roundedCorners: true,
      chartBorderColor: '0055AA',
      chartBorderWidthEmu: 25_400,
      chartBorderCompound: 'dbl',
    }), RECT, 1);

    expect(rec.clipCalls).toBe(1);
    // Four clip corners and four corners for each of the two border rails.
    expect(rec.quadratics).toHaveLength(12);
  });

  it('keeps an explicit false rectangular and preserves rounded noFill clipping', () => {
    const sharp = recordingCtx();
    renderChart(sharp.ctx, baseModel({
      roundedCorners: false,
      chartBg: 'F2F2F2',
      chartBorderColor: '0055AA',
    }), RECT, 1);
    expect(sharp.clipCalls).toBe(0);
    expect(sharp.strokeRects).toHaveLength(1);
    expect(sharp.quadratics).toHaveLength(0);

    const noFill = recordingCtx();
    renderChart(noFill.ctx, baseModel({
      roundedCorners: true,
      chartBg: null,
      chartBorderColor: '0055AA',
    }), RECT, 1);
    expect(noFill.clipCalls).toBe(1);
    expect(noFill.rects).toHaveLength(0);
    expect(noFill.quadratics).toHaveLength(8);
  });

  it('uses the shared gradient recipe inside the rounded clip and honors host rotation', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      roundedCorners: true,
      chartFill: {
        fillType: 'gradient',
        stops: [
          { position: 0, color: '112233' },
          { position: 1, color: 'AABBCC' },
        ],
        angle: 90,
        gradType: 'linear',
        rotWithShape: false,
      },
    }), RECT, 1, 30);

    expect(rec.clipCalls).toBe(1);
    expect(rec.gradients).toHaveLength(1);
    const [x1, y1, x2, y2] = rec.gradients[0].args;
    expect((y2 - y1) / (x2 - x1)).toBeCloseTo(Math.sqrt(3), 5);
    expect(rec.gradients[0].stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(170,187,204,1)' },
    ]);
  });

  it('uses the shared pattern paint inside the rounded clip', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      roundedCorners: true,
      chartFill: {
        fillType: 'pattern', fg: '112233', bg: 'AABBCC', preset: 'diagCross',
      },
    }), RECT, 1);

    expect(rec.clipCalls).toBe(1);
    expect(rec.rects[0]).toMatchObject({
      x: 0, y: 0, w: 640, h: 360, fs: 'rgba(17,34,51,1)',
    });
  });

  it('lets linked chart-area paint replace only an unauthored host default', () => {
    const chart = baseModel({
      chartBg: 'FFFFFF',
      chartStyleRoles: {
        chartArea: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          lineColors: ['445566'],
          lineWidthEmu: 9525,
          lineCustomDash: [{ dash: 1.25, space: 0.75 }],
          lineCompound: 'dbl',
          lineCap: 'sq',
          lineJoin: 'bevel',
        },
      },
    });
    const linked = recordingCtx();
    renderChart(linked.ctx, chart, RECT, 1);
    expect(linked.gradients).toHaveLength(1);
    expect(linked.rects[0]).toMatchObject({
      x: 0, y: 0, w: 640, h: 360, fs: '[object Object]',
    });
    expect(linked.strokeRects.filter(rect => rect.ss === '#445566')).toEqual([
      expect.objectContaining({ lw: 0.25, cap: 'square', join: 'bevel' }),
      expect.objectContaining({ lw: 0.25, cap: 'square', join: 'bevel' }),
    ]);
    expect(linked.strokeRects.find(rect => rect.ss === '#445566')?.dash)
      .toEqual([0.9375, 0.5625]);

    const directEmptyDash = recordingCtx();
    renderChart(directEmptyDash.ctx, { ...chart, chartBorderCustomDash: [] }, RECT, 1);
    expect(directEmptyDash.strokeRects.find(rect => rect.ss === '#445566')?.dash)
      .toEqual([]);

    const directNoFill = recordingCtx();
    renderChart(directNoFill.ctx, {
      ...chart,
      chartBg: null,
      chartFillHidden: true,
      chartFillPaintAuthored: true,
      chartBorderHidden: true,
      chartBorderPaintAuthored: true,
    }, RECT, 1);
    expect(directNoFill.gradients).toHaveLength(0);
    expect(directNoFill.rects).toHaveLength(0);
    expect(directNoFill.strokeRects).toHaveLength(0);
  });
});

describe('rich chart titles', () => {
  it('wraps a long heading and preserves an explicit italic subtitle line', () => {
    const rec = recordingCtx((text, fontPx) => text.length * fontPx * 0.6);
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      title: 'A long heading that must wrap\nSubtitle',
      titleFontSizeHpt: 1800,
      titleRichRuns: [
        { text: 'A long heading that must wrap', fontSizeHpt: 1800, bold: true },
        { text: '\nSubtitle', fontSizeHpt: 1400, italic: true, color: '112233' },
      ],
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
    }), { x: 0, y: 0, w: 180, h: 220 }, 1);

    const titlePieces = rec.texts.filter(text =>
      ['A', 'long', 'heading', 'that', 'must', 'wrap', 'Subtitle'].includes(text.text.trim()),
    );
    expect(new Set(titlePieces.map(text => text.y)).size).toBeGreaterThanOrEqual(3);
    const subtitle = rec.texts.find(text => text.text === 'Subtitle');
    expect(subtitle?.font).toContain('italic');
    expect(subtitle?.font).toContain('14px');
    expect(subtitle?.fillStyle).toBe('#112233');
  });
});

describe('classic 3-D compatibility projection', () => {
  it('uses the canonical 2-D family fallback when the optional renderer is absent', () => {
    const rec = recordingCtx();
    renderChartCore(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({ values: [5] })],
    }), RECT, 1, 0);
    expect(rec.rects).toHaveLength(1);
    expect(materialFills(rec, '4472C4')).toHaveLength(0);
  });

  it('does not apply the mesh face budget to the tree-shaken 2-D fallback', () => {
    const rec = recordingCtx();
    const categories = Array.from({ length: 300 }, (_, index) => String(index));
    renderChartCore(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories,
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({ values: new Array(300).fill(5) })],
    }), RECT, 1, 0);
    expect(rec.texts.map(item => item.text)).not.toContain('(too many data points)');
    expect(rec.rects).toHaveLength(300);
  });

  it.each([
    ['clusteredBar', false],
    ['stackedBar', false],
    ['clusteredBarH', false],
    ['line', false],
    ['area', false],
    ['pie', true],
  ] as const)('routes %s through the shared 3-D painter', (chartType, radial) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      title: `3D ${chartType}`,
      categories: ['A', 'B', 'C'],
      showLegend: true,
      showDataLabels: radial,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        gapDepthPercent: 150,
      },
      series: [
        series({ name: 'First', values: [2, 4, 6], categories: ['A', 'B', 'C'] }),
        series({ name: 'Second', values: [3, 5, 7], categories: ['A', 'B', 'C'] }),
      ],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toContain(`3D ${chartType}`);
    expect(rec.texts.map(text => text.text)).not.toContain(`Chart: ${chartType}`);
    expect(rec.paintEvents.filter(event => event.kind === 'stroke').length).toBeGreaterThan(0);
    if (radial) {
      expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining(['2', '4', '6']));
    } else if (chartType.includes('Bar')) {
      // 3-D bars are projected polygon faces. A fillRect would bypass the
      // shared projection and make bar slopes disagree with axes/gridlines.
      expect(
        materialFills(rec, '4472C4').length + materialFills(rec, 'ED7D31').length,
      ).toBeGreaterThanOrEqual(6);
    }
  });

  it('uses the shared automatic tick density with only an authored 3-D maximum', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      valMax: 8,
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        gapDepthPercent: 150,
      },
      series: [series({ values: [8, 6, 4, 2] })],
    }), { x: 0, y: 0, w: 300, h: 190 }, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining(['0', '1', '2', '3']));
  });

  it('keeps the ordinary one-sided tick density when 3-D display units are authored', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      valMax: 8,
      valAxisDisplayUnits: { divisor: 100, builtInUnit: 'hundreds', label: null },
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        gapDepthPercent: 150,
      },
      series: [series({ values: [8, 6, 4, 2] })],
    }), { x: 0, y: 0, w: 300, h: 190 }, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining(['0', '0.01', '0.02', '0.03']));
  });

  it('keeps the ordinary 2-D bar path when view3D is absent', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Category'],
      series: [series({ values: [5] })],
    }), RECT, 1);
    expect(rec.rects).toHaveLength(1);
  });

  it('fans default grid lines toward one finite perspective vanishing point', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: true,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [10, 20] })],
    }), RECT, 1);
    const wallLines = rec.strokes
      // Style 2's automatic 3-D axis/grid stroke resolves to the same neutral
      // gray used by Office's vector output; do not keep the older #A6A6A6
      // compatibility color in this geometry characterization.
      .filter(stroke => stroke.ss === '#898989' && stroke.points.length === 2)
      .map(stroke => ({
        x0: stroke.points[0].x, y0: stroke.points[0].y,
        x1: stroke.points[1].x, y1: stroke.points[1].y,
      }))
      .filter(segment => segment.x1 - segment.x0 > 100);
    const slopes = wallLines.map(segment =>
      (segment.y1 - segment.y0) / (segment.x1 - segment.x0)
    );
    expect(slopes.length).toBeGreaterThan(3);
    expect(Math.min(...slopes)).toBeGreaterThan(0.04);
    expect(Math.max(...slopes)).toBeLessThan(0.16);
    expect(Math.max(...slopes) - Math.min(...slopes)).toBeGreaterThan(0.02);
  });

  it('does not invent 3-D value gridlines when majorGridlines is omitted', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisGridlineColor: 'FF00FF',
      valMin: 0,
      valMax: 20,
      valAxisMajorUnit: 5,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);

    expect(rec.strokes.filter(stroke => stroke.ss === '#FF00FF')).toHaveLength(0);
  });

  it('does not invent 3-D category-depth gridlines when majorGridlines is omitted', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      catAxisLineColor: '000000',
      valAxisLineColor: '000000',
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        floor: { lineHidden: true },
        sideWall: { lineHidden: true },
        backWall: { lineHidden: true },
        seriesAxis: { hidden: true, lineHidden: true, majorTickMark: 'none' },
      },
      series: [series({ values: [5, 15], lineHidden: true })],
    }), RECT, 1);

    expect(rec.strokes.filter(stroke => stroke.ss === '#898989')).toHaveLength(0);
  });

  it.each([0.5, 1, 2])(
    'uses the same authored 0.25pt width for 3-D axis rules and ticks at %sx',
    (ptToPx) => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 20,
      valAxisMajorUnit: 10,
      catAxisLineColor: 'FF00FF',
      catAxisLineWidthEmu: 3_175,
      catAxisMajorTickMark: 'out',
      catAxisMajorGridlines: false,
      valAxisLineColor: '00AA00',
      valAxisLineWidthEmu: 3_175,
      valAxisMajorTickMark: 'out',
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        barGrouping: 'standard',
        floor: { lineHidden: true },
        sideWall: { lineHidden: true },
        backWall: { lineHidden: true },
        seriesAxis: {
          hidden: false,
          lineHidden: false,
          lineColor: '0000FF',
          lineWidthEmu: 3_175,
          majorTickMark: 'out',
        },
      },
      series: [
        series({ values: [5, 15], lineHidden: true }),
        series({ values: [8, 12], lineHidden: true }),
      ],
    }), RECT, ptToPx);

    for (const color of ['#FF00FF', '#00AA00', '#0000FF']) {
      const strokes = rec.strokes.filter(stroke => stroke.ss === color);
      expect(strokes.length).toBeGreaterThan(0);
      const length = (stroke: (typeof strokes)[number]) => Math.hypot(
        stroke.points[1].x - stroke.points[0].x,
        stroke.points[1].y - stroke.points[0].y,
      );
      const frame = strokes.reduce((longest, stroke) =>
        length(stroke) > length(longest) ? stroke : longest);
      const ticks = strokes.filter(stroke => stroke !== frame);
      expect(frame.lw).toBe(0.25 * ptToPx);
      expect(ticks.length).toBeGreaterThan(0);
      expect(ticks.every(stroke => stroke.lw === 0.25 * ptToPx)).toBe(true);
    }
  });

  it('gives visible axes ownership of coincident 3-D surface edges', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 20,
      valAxisMajorGridlines: false,
      catAxisLineColor: 'FF00FF',
      catAxisLineWidthEmu: 3_175,
      catAxisLinePaintAuthored: true,
      valAxisLineColor: '00AA00',
      valAxisLineWidthEmu: 3_175,
      valAxisLinePaintAuthored: true,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        floor: { lineColor: 'FF0000', lineWidthEmu: 12_700 },
        sideWall: { lineColor: '0000FF', lineWidthEmu: 12_700 },
        backWall: { lineColor: '808080', lineWidthEmu: 12_700 },
      },
      series: [series({ values: [5, 15], lineHidden: true })],
    }), RECT, 1);

    const samePoint = (
      left: { x: number; y: number }, right: { x: number; y: number },
    ) => Math.hypot(left.x - right.x, left.y - right.y) < 1e-6;
    const ownsEdge = (
      axis: (typeof rec.strokes)[number], surface: (typeof rec.strokes)[number],
    ) => surface.points.some((point, index) => {
      const next = surface.points[(index + 1) % surface.points.length];
      return (samePoint(axis.points[0], point) && samePoint(axis.points[1], next))
        || (samePoint(axis.points[0], next) && samePoint(axis.points[1], point));
    });
    const axes = rec.strokes.filter(stroke =>
      (stroke.ss === '#FF00FF' || stroke.ss === '#00AA00')
      && stroke.points.length === 2
    );
    const surfaces = rec.strokes.filter(stroke =>
      ['#FF0000', '#0000FF', '#808080'].includes(stroke.ss)
    );

    expect(axes.length).toBeGreaterThanOrEqual(2);
    expect(axes.every(axis => surfaces.every(surface => !ownsEdge(axis, surface)))).toBe(true);
  });

  it('places standard 3-D category and series ticks on slot boundaries', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      catAxisCrossBetween: 'between',
      catAxisLineColor: 'FF00FF',
      catAxisMajorTickMark: 'out',
      catAxisMajorGridlines: false,
      valAxisHidden: true,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 0,
        rightAngleAxes: true,
        barGrouping: 'standard',
        floor: { lineHidden: true },
        sideWall: { lineHidden: true },
        backWall: { lineHidden: true },
        seriesAxis: {
          hidden: false,
          lineHidden: false,
          lineColor: '0000FF',
          majorTickMark: 'out',
        },
      },
      series: [
        series({ name: 'North', values: [5, 10, 15], lineHidden: true }),
        series({ name: 'South', values: [7, 12, 17], lineHidden: true }),
      ],
    }), RECT, 1);

    const tickFractions = (color: string) => {
      const strokes = rec.strokes.filter(stroke => stroke.ss === color && stroke.points.length === 2);
      const length = (stroke: (typeof strokes)[number]) => Math.hypot(
        stroke.points[1].x - stroke.points[0].x,
        stroke.points[1].y - stroke.points[0].y,
      );
      const axis = strokes.reduce((longest, stroke) =>
        length(stroke) > length(longest) ? stroke : longest);
      const [start, end] = axis.points;
      const dx = end.x - start.x;
      const dy = end.y - start.y;
      const lengthSquared = dx * dx + dy * dy;
      return strokes
        .filter(stroke => stroke !== axis)
        .map(stroke => {
          const anchor = stroke.points[1];
          return ((anchor.x - start.x) * dx + (anchor.y - start.y) * dy) / lengthSquared;
        })
        .sort((a, b) => a - b);
    };

    expect(tickFractions('#FF00FF')).toEqual([
      expect.closeTo(0, 6), expect.closeTo(1 / 3, 6),
      expect.closeTo(2 / 3, 6), expect.closeTo(1, 6),
    ]);
    expect(tickFractions('#0000FF')).toEqual([
      expect.closeTo(0, 6), expect.closeTo(0.5, 6), expect.closeTo(1, 6),
    ]);
  });

  it('does not turn a linked gridline style into unauthored 3-D grid geometry', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      chartStyleRoles: {
        gridlineMajor: { lineColors: ['FF00FF'], lineWidthEmu: 12_700 },
      },
      valAxisMajorGridlines: undefined,
      catAxisMajorGridlines: undefined,
      catAxisLineColor: '000000',
      valAxisLineColor: '000000',
      threeD: {
        rotationX: 15,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        floor: { lineColor: '808080', lineWidthEmu: 12_700 },
        sideWall: { lineColor: '808080', lineWidthEmu: 12_700 },
        backWall: { lineColor: '808080', lineWidthEmu: 12_700 },
        seriesAxis: { hidden: true, lineHidden: true, majorTickMark: 'none' },
      },
      series: [series({ values: [5, 15], lineHidden: true })],
    }), RECT, 1);

    expect(rec.strokes.filter(stroke => stroke.ss === '#FF00FF')).toHaveLength(0);
    const wallStrokes = rec.strokes.filter(stroke => stroke.ss === '#808080');
    expect(wallStrokes.length).toBeGreaterThanOrEqual(3);
    expect(wallStrokes.every(stroke => stroke.points.length >= 2)).toBe(true);
  });

  it('draws 6pt major and 4pt minor value ticks as horizontal screen annotations', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 20,
      valAxisMajorUnit: 10,
      valAxisMinorUnit: 2,
      valAxisMajorTickMark: 'out',
      valAxisMinorTickMark: 'out',
      valAxisMajorGridlines: false,
      catAxisMajorTickMark: 'none',
      valAxisLineColor: 'FF00FF',
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    const axisStrokes = rec.strokes.filter(stroke => stroke.ss === '#FF00FF' && stroke.points.length === 2);
    const length = (stroke: (typeof axisStrokes)[number]) => Math.hypot(
      stroke.points[1].x - stroke.points[0].x,
      stroke.points[1].y - stroke.points[0].y,
    );
    const axis = axisStrokes.reduce((longest, stroke) => length(stroke) > length(longest) ? stroke : longest);
    const ticks = axisStrokes.filter(stroke => stroke !== axis);
    const lengths = ticks.map(length);
    expect(lengths.some(value => Math.abs(value - 6) < 1e-8)).toBe(true);
    expect(lengths.some(value => Math.abs(value - 4) < 1e-8)).toBe(true);
    for (const tick of ticks) {
      const tickVector = {
        x: tick.points[1].x - tick.points[0].x,
        y: tick.points[1].y - tick.points[0].y,
      };
      expect(Math.abs(tickVector.y)).toBeLessThan(1e-8);
    }
  });

  it('places 3-D value and category labels beyond their outward major ticks', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      valAxisMajorTickMark: 'out',
      catAxisMajorTickMark: 'out',
      valAxisLineColor: 'FF00FF',
      catAxisLineColor: '00AA00',
      valAxisMajorGridlines: false,
      catAxisLabelRotation: 0,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [5, 8] })],
    }), RECT, 1);

    const valueZero = rec.texts.find(text => text.text === '0');
    const categoryA = rec.texts.find(text => text.text === 'A');
    expect(valueZero).toBeDefined();
    expect(categoryA).toBeDefined();

    const valueTick = rec.strokes.find(stroke =>
      stroke.ss === '#FF00FF'
      && stroke.points.length === 2
      && Math.abs(Math.abs(stroke.points[1].x - stroke.points[0].x) - 6) < 0.001
      && Math.abs(stroke.points[1].y - stroke.points[0].y) < 0.001
      && Math.abs(stroke.points[0].y - valueZero!.y) < 0.001);
    const categoryTick = rec.strokes.find(stroke =>
      stroke.ss === '#00AA00'
      && stroke.points.length === 2
      && Math.abs(Math.abs(stroke.points[1].y - stroke.points[0].y) - 6) < 0.001
      && Math.abs(stroke.points[1].x - stroke.points[0].x) < 0.001
      && Math.abs(stroke.points[0].x - categoryA!.x) < 0.001);
    expect(valueTick).toBeDefined();
    expect(categoryTick).toBeDefined();

    if (valueZero!.align === 'right') {
      expect(valueZero!.x).toBeLessThan(Math.min(valueTick!.points[0].x, valueTick!.points[1].x));
    } else {
      expect(valueZero!.x).toBeGreaterThan(Math.max(valueTick!.points[0].x, valueTick!.points[1].x));
    }
    if (categoryA!.baseline === 'top') {
      expect(categoryA!.y).toBeGreaterThan(Math.max(categoryTick!.points[0].y, categoryTick!.points[1].y));
    } else {
      expect(categoryA!.y).toBeLessThan(Math.min(categoryTick!.points[0].y, categoryTick!.points[1].y));
    }
  });

  it('does not invent 3-D minor ticks when minorTickMark is omitted', () => {
    const minorTickCount = (minorTickMark: ChartModel['valAxisMinorTickMark']): number => {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line', categories: ['A', 'B'],
        valMin: 0, valMax: 20, valAxisMajorUnit: 10,
        valAxisMinorTickMark: minorTickMark,
        valAxisMajorGridlines: false,
        catAxisMajorTickMark: 'none',
        valAxisLineColor: 'FF00FF',
        threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
        series: [series({ values: [5, 15] })],
      }), RECT, 1);
      return rec.strokes
        .filter(stroke => stroke.ss === '#FF00FF' && stroke.points.length === 2)
        .map(stroke => Math.hypot(
          stroke.points[1].x - stroke.points[0].x,
          stroke.points[1].y - stroke.points[0].y,
        ))
        .filter(length => Math.abs(length - 4) < 1e-8)
        .length;
    };

    expect(minorTickCount(undefined)).toBe(0);
    expect(minorTickCount('none')).toBe(0);
    expect(minorTickCount('cross')).toBe(8);
  });

  it('keeps the horizontal 3-D value grid non-degenerate between floor and back wall', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      valAxisMajorGridlines: true,
      valAxisGridlineColor: 'FF0000',
      valMin: 0,
      valMax: 20,
      valAxisMajorUnit: 10,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    const grid = rec.strokes.filter(stroke => stroke.ss === '#FF0000');
    expect(grid.length).toBeGreaterThanOrEqual(3);
    expect(grid.every(stroke => Math.hypot(
      stroke.points.at(-1)!.x - stroke.points[0].x,
      stroke.points.at(-1)!.y - stroke.points[0].y,
    ) > 1)).toBe(true);
  });

  it('uses the projected 3-D value-axis length for an automatic explicit-span unit', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 5,
      valAxisFormatCode: '0',
      valAxisMajorGridlines: false,
      catAxisMajorTickMark: 'none',
      threeD: {
        rotationX: 20,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
      },
      series: [series({ values: [2, 5], showMarker: false })],
    }), { x: 0, y: 0, w: 738, h: 439 }, 4 / 3);
    expect(rec.texts.map(item => item.text).filter(text => /^\d+$/.test(text)))
      .toEqual(['0', '1', '2', '3', '4', '5']);
  });

  it('keeps unfilled 3-D floor and walls transparent while honoring authored fills', () => {
    const automatic = recordingCtx();
    renderChart(automatic.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);

    const explicitNoFill = recordingCtx();
    renderChart(explicitNoFill.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
        floor: { fillHidden: true },
        sideWall: { fillHidden: true },
        backWall: { fillHidden: true },
      },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    expect(automatic.filledPaths).toEqual(explicitNoFill.filledPaths);

    const authored = recordingCtx();
    renderChart(authored.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
        backWall: { fillColor: 'F2F2F2', fillHidden: false },
      },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    expect(authored.filledPaths.some(path => path.fillStyle === '#F2F2F2')).toBe(true);
  });

  it('paints authored CT_Surface thickness as projected slabs without leaving the chart bounds', () => {
    const render = (thicknessPercent: number | undefined) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A', 'B', 'C'],
        valAxisMajorGridlines: false,
        threeD: {
          rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
          floor: { fillColor: 'FF0000', thicknessPercent },
          sideWall: { fillColor: '00B050', thicknessPercent },
          backWall: { fillColor: 'AA00FF', thicknessPercent },
        },
        series: [series({ values: [2, 8, 4], showMarker: false })],
      }), RECT, 1);
      return rec;
    };
    const omitted = render(undefined);
    const planar = render(0);
    const thick = render(25);
    expect(planar.filledPaths).toEqual(omitted.filledPaths);
    expect(planar.strokedPaths).toEqual(omitted.strokedPaths);
    for (const color of ['#FF0000', '#00B050', '#AA00FF']) {
      expect(planar.filledPaths.filter(path => path.fillStyle === color)).toHaveLength(1);
      expect(thick.filledPaths.filter(path => path.fillStyle === color).length).toBeGreaterThan(1);
    }
    const points = thick.filledPaths.flatMap(path => path.points);
    expect(Math.min(...points.map(point => point.x))).toBeGreaterThanOrEqual(RECT.x - 1e-9);
    expect(Math.max(...points.map(point => point.x))).toBeLessThanOrEqual(RECT.x + RECT.w + 1e-9);
    expect(Math.min(...points.map(point => point.y))).toBeGreaterThanOrEqual(RECT.y - 1e-9);
    expect(Math.max(...points.map(point => point.y))).toBeLessThanOrEqual(RECT.y + RECT.h + 1e-9);
  });

  it('continues authored category and value gridlines across visible thick CT_Surface faces', () => {
    const render = (thicknessPercent: number, chartType = 'line') => {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        catAxisCrossBetween: 'between',
        catAxisMajorGridlines: true,
        catAxisGridlineColor: 'FF00FF',
        catAxisGridlineWidthEmu: 25_400,
        catAxisGridlineDash: 'dash',
        catAxisMinorGridlines: true,
        catAxisMinorGridlineColor: 'FF8000',
        catAxisMinorGridlineWidthEmu: 25_400,
        catAxisMinorGridlineDash: 'dot',
        valAxisMajorGridlines: true,
        valAxisGridlineColor: '00FFFF',
        valAxisGridlineWidthEmu: 25_400,
        valAxisGridlineDash: 'dot',
        valAxisMinorGridlines: true,
        valAxisMinorGridlineColor: '123456',
        valAxisMinorGridlineWidthEmu: 25_400,
        valAxisMinorGridlineDash: 'dash',
        valAxisMinorUnit: 1,
        valMin: 0,
        valMax: 10,
        valAxisMajorUnit: 2,
        threeD: {
          rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
          floor: { fillColor: 'C00000', thicknessPercent },
          sideWall: { fillColor: '008000', thicknessPercent },
          backWall: { fillColor: '4472C4', thicknessPercent },
        },
        series: [series({ values: [2, 8, 4], showMarker: false })],
      }), RECT, 1);
      return rec.strokes;
    };
    const planar = render(0);
    const thick = render(25);
    for (const color of ['#FF00FF', '#FF8000', '#00FFFF', '#123456']) {
      const planarLines = planar.filter(stroke => stroke.ss === color);
      const thickLines = thick.filter(stroke => stroke.ss === color);
      expect(planarLines.length).toBeGreaterThan(0);
      expect(thickLines.length).toBeGreaterThan(0);
      expect(thickLines.length).toBeGreaterThan(planarLines.length);
      expect(thickLines.every(stroke => stroke.lw === 2)).toBe(true);
      expect(thickLines.every(stroke => stroke.dash.length > 0)).toBe(true);
    }
    for (const chartType of ['area', 'clusteredBar', 'clusteredBarH']) {
      const planarFamily = render(0, chartType);
      const thickFamily = render(25, chartType);
      for (const color of ['#FF00FF', '#FF8000', '#00FFFF', '#123456']) {
        expect(planarFamily.filter(stroke => stroke.ss === color).length)
          .toBeGreaterThan(0);
        expect(thickFamily.filter(stroke => stroke.ss === color).length)
          .toBeGreaterThan(planarFamily.filter(stroke => stroke.ss === color).length);
      }
    }
  });

  it('prefetches and projects only applicable flat CT_Surface stretch pictures', () => {
    const picture = {
      fillType: 'image' as const,
      imagePath: 'xl/media/wall.png',
      mimeType: 'image/png',
      stretch: true,
    };
    const model = (applyToFront: boolean) => baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 10,
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
        backWall: {
          thicknessPercent: 0,
          style: { fillPaints: [picture], fillPaintAuthored: true },
          pictureOptions: { applyToFront, pictureFormat: 'stretch' },
        },
      },
      series: [series({ values: [2, 8, 4], showMarker: false })],
    });
    const bitmap = { width: 80, height: 40 } as unknown as CanvasImageSource;
    expect(collectChartMarkerImageFills(model(true))).toEqual([picture]);
    expect(collectChartMarkerImageFills(model(false))).toEqual([]);

    const visible = recordingCtx();
    renderChartCore(visible.ctx, model(true), RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(visible.drawImages.length).toBeGreaterThan(0);
    const hidden = recordingCtx();
    renderChartCore(hidden.ctx, model(false), RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(hidden.drawImages).toHaveLength(0);

    const linked = baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 10,
      valAxisMajorGridlines: false,
      threeD: { rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30 },
      chartStyleRoles: {
        wall: { fillPaints: [picture], fillPaintAuthored: true },
      },
      series: [series({ values: [2, 8], showMarker: false })],
    });
    expect(collectChartMarkerImageFills(linked)).toEqual([picture]);
    const linkedRec = recordingCtx();
    renderChartCore(linkedRec.ctx, linked, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(linkedRec.drawImages.length).toBeGreaterThan(0);
  });

  it.each(['line', 'surface3D'] as const)(
    'projects independently selected positive-thickness CT_Surface stretch faces for %s',
    chartType => {
      const picture = {
        fillType: 'image' as const,
        imagePath: 'xl/media/thick-surface.png',
        mimeType: 'image/png',
        stretch: true,
      };
      const model = (
        kind: 'floor' | 'sideWall' | 'backWall',
        pictureOptions: ChartThreeDPictureOptions,
      ) => baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        valMin: 0,
        valMax: 10,
        valAxisMajorGridlines: false,
        threeD: {
          rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
          [kind]: {
            thicknessPercent: 25,
            lineColor: 'FF0000',
            style: { fillPaints: [picture], fillPaintAuthored: true },
            pictureOptions: { ...pictureOptions, pictureFormat: 'stretch' },
          },
        },
        series: chartType === 'surface3D'
          ? [series({ values: [2, 8, 4] }), series({ values: [4, 6, 7] })]
          : [series({ values: [2, 8, 4], showMarker: false })],
      });
      const bitmap = { width: 80, height: 40 } as unknown as CanvasImageSource;
      const draws = (
        kind: 'floor' | 'sideWall' | 'backWall',
        pictureOptions: ChartThreeDPictureOptions,
        reversed = false,
      ) => {
        const chart = {
          ...model(kind, pictureOptions),
          valAxisOrientation: reversed ? 'maxMin' as const : undefined,
        };
        expect(collectChartMarkerImageFills(chart)).toEqual([picture]);
        const rec = recordingCtx();
        renderChartCore(rec.ctx, chart, RECT, 1, 0, testThreeD, undefined, () => bitmap);
        expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
        return rec.drawImages.length;
      };
      for (const [kind, pictureOptions] of [
        ['backWall', { applyToFront: true, applyToSides: false, applyToEnd: false }],
        ['sideWall', { applyToFront: false, applyToSides: true, applyToEnd: false }],
        ['floor', { applyToFront: false, applyToSides: false, applyToEnd: true }],
      ] as const) {
        const normal = draws(kind, pictureOptions);
        expect(normal, kind).toBeGreaterThan(0);
        expect(draws(kind, pictureOptions, true), `${kind} reversed`).toBe(normal);
      }
      expect(collectChartMarkerImageFills(model('backWall', {
        applyToFront: false, applyToSides: false, applyToEnd: false,
      }))).toEqual([]);
    },
  );

  it('repeats stackScale pictures by value units but keeps the Office floor exception', () => {
    const picture = {
      fillType: 'image' as const,
      imagePath: 'xl/media/stack-scale.png',
      mimeType: 'image/png',
      stretch: true,
    };
    const model = (
      kind: 'backWall' | 'sideWall' | 'floor',
      pictureFormat: 'stretch' | 'stackScale' | 'stack',
      thicknessPercent = 0,
    ) => baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 10,
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
        [kind]: {
          thicknessPercent,
          style: { fillPaints: [picture], fillPaintAuthored: true },
          pictureOptions: {
            ...(kind === 'backWall' ? { applyToFront: true } : { applyToSides: true }),
            pictureFormat,
            pictureStackUnit: pictureFormat === 'stackScale' ? 2 : undefined,
          },
        },
      },
      series: [series({ values: [2, 8, 4], showMarker: false })],
    });
    const bitmap = { width: 80, height: 40 } as unknown as CanvasImageSource;
    const draws = (chart: ChartModel) => {
      const rec = recordingCtx();
      renderChartCore(rec.ctx, chart, RECT, 1, 0, testThreeD, undefined, () => bitmap);
      return rec.drawImages.length;
    };
    const backStretch = draws(model('backWall', 'stretch'));
    const backStack = draws(model('backWall', 'stack'));
    const backStackScale = draws(model('backWall', 'stackScale'));
    expect(backStretch).toBeGreaterThan(0);
    expect(backStack).toBeGreaterThan(0);
    expect(backStackScale).toBeGreaterThan(backStretch);
    expect(draws(model('floor', 'stackScale'))).toBe(draws(model('floor', 'stretch')));
    expect(collectChartMarkerImageFills(model('backWall', 'stack'))).toEqual([picture]);
    const reversedStack = {
      ...model('backWall', 'stack'),
      valAxisOrientation: 'maxMin' as const,
    } satisfies ChartModel;
    expect(collectChartMarkerImageFills(reversedStack)).toEqual([picture]);
    expect(draws(reversedStack)).toBe(backStack);
    const thickStack = model('backWall', 'stack', 25);
    expect(collectChartMarkerImageFills(thickStack)).toEqual([picture]);
    expect(draws(thickStack)).toBeGreaterThan(0);
    expect(collectChartMarkerImageFills(model('backWall', 'stretch', 25))).toEqual([picture]);
    for (const kind of ['backWall', 'sideWall'] as const) {
      const thickStretch = model(kind, 'stretch', 25);
      const thickStackScale = model(kind, 'stackScale', 25);
      expect(collectChartMarkerImageFills(thickStackScale)).toEqual([picture]);
      expect(draws(thickStackScale)).toBeGreaterThan(draws(thickStretch));
    }
    const thickFloorStretch = model('floor', 'stretch', 25);
    const thickFloorStackScale = model('floor', 'stackScale', 25);
    expect(collectChartMarkerImageFills(thickFloorStackScale)).toEqual([picture]);
    expect(draws(thickFloorStackScale)).toBe(draws(thickFloorStretch));
    const reversedStretch = {
      ...model('backWall', 'stretch'),
      valAxisOrientation: 'maxMin',
    } satisfies ChartModel;
    const reversedStackScale = {
      ...model('backWall', 'stackScale'),
      valAxisOrientation: 'maxMin',
    } satisfies ChartModel;
    expect(collectChartMarkerImageFills(reversedStretch)).toEqual([picture]);
    expect(collectChartMarkerImageFills(reversedStackScale)).toEqual([picture]);
    expect(draws(reversedStretch)).toBe(backStretch);
    expect(draws(reversedStackScale)).toBe(backStackScale);
    const invalidProvenance = model('backWall', 'stretch');
    invalidProvenance.threeD!.backWall!.pictureOptions = {
      applyToFront: true,
      pictureFormatAuthored: true,
      pictureStackUnitAuthored: true,
    };
    expect(collectChartMarkerImageFills(invalidProvenance)).toEqual([]);
    const croppedPicture = {
      ...picture,
      srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
    };
    const cropped = model('backWall', 'stretch');
    cropped.threeD!.backWall!.style!.fillPaints = [croppedPicture];
    expect(collectChartMarkerImageFills(cropped)).toEqual([croppedPicture]);
    const croppedRec = recordingCtx();
    renderChartCore(croppedRec.ctx, cropped, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    const croppedSourceXs = croppedRec.drawImages
      .filter(call => call[0] === bitmap)
      .map(call => Number(call[1]));
    expect(croppedSourceXs.length).toBeGreaterThan(0);
    expect(Math.min(...croppedSourceXs)).toBeGreaterThan(15);
    expect(collectChartMarkerImageFills({
      ...model('backWall', 'stackScale'),
      valMax: 8_192,
    })).toEqual([picture]);
    expect(collectChartMarkerImageFills({
      ...model('backWall', 'stackScale'),
      valMax: 8_194,
    })).toEqual([]);
  });

  it('projects bounded DrawingML tile grids across planar and thick Surface faces', () => {
    const tiled = {
      fillType: 'image' as const,
      imagePath: 'xl/media/surface-tile.png',
      mimeType: 'image/png',
      stretch: false,
      dpi: 96,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    const model = (
      thicknessPercent: number,
      scale = 1,
      srcRect?: { l: number; t: number; r: number; b: number },
    ) => baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 10,
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 20,
        rotationY: 20,
        depthPercent: 100,
        perspective: 30,
        backWall: {
          thicknessPercent,
          style: {
            fillPaints: [{
              ...tiled,
              srcRect,
              tile: { ...tiled.tile, sx: scale, sy: scale },
            }],
            fillPaintAuthored: true,
          },
          pictureOptions: {
            applyToFront: true,
            applyToSides: true,
            applyToEnd: true,
            pictureFormat: 'stretch',
          },
        },
      },
      series: [series({ values: [2, 8, 4], showMarker: false })],
    });
    const bitmap = { width: 80, height: 40 } as unknown as CanvasImageSource;
    const draws = (chart: ChartModel) => {
      expect(collectChartMarkerImageFills(chart)).toHaveLength(1);
      const rec = recordingCtx();
      class TestOffscreenCanvas {
        readonly width: number;
        readonly height: number;
        constructor(width: number, height: number) {
          this.width = width;
          this.height = height;
        }
        getContext(): CanvasRenderingContext2D { return rec.ctx; }
      }
      vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);
      try {
        renderChartCore(rec.ctx, chart, RECT, 1, 0, testThreeD, undefined, () => bitmap);
        return rec.drawImages.length;
      } finally {
        vi.unstubAllGlobals();
      }
    };
    const planar = draws(model(0));
    expect(planar).toBeGreaterThan(0);
    expect(draws(model(0, 0.5))).toBeGreaterThan(planar);
    expect(draws(model(25))).toBeGreaterThan(0);
    expect(draws(model(0, 1, { l: 0.25, t: 0, r: 0, b: 0 }))).toBeGreaterThan(0);
    expect(draws(model(25, 1, { l: -0.25, t: 0, r: 0, b: 0 }))).toBeGreaterThan(0);
  });

  it('uses the same planar CT_Surface picture path for Surface3D', () => {
    const picture = {
      fillType: 'image' as const,
      imagePath: 'xl/media/surface-wall.png',
      mimeType: 'image/png',
      stretch: true,
    };
    const model = baseModel({
      chartType: 'surface3D',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 2,
      threeD: {
        rotationX: 20, rotationY: 20, depthPercent: 100, perspective: 30,
        sideWall: {
          thicknessPercent: 0,
          lineColor: 'FF0000',
          style: { fillPaints: [picture], fillPaintAuthored: true },
          pictureOptions: {
            applyToSides: true, pictureFormat: 'stackScale', pictureStackUnit: 2,
          },
        },
      },
      series: [
        series({ values: [2, 8] }),
        series({ values: [4, 6] }),
      ],
    });
    const bitmap = { width: 80, height: 40 } as unknown as CanvasImageSource;
    expect(collectChartMarkerImageFills(model)).toEqual([picture]);
    const rec = recordingCtx();
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(rec.drawImages.length).toBeGreaterThan(1);
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
    const reversed = { ...model, valAxisOrientation: 'maxMin' as const };
    expect(collectChartMarkerImageFills(reversed)).toEqual([picture]);
    const reversedRec = recordingCtx();
    renderChartCore(reversedRec.ctx, reversed, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(reversedRec.drawImages.length).toBe(rec.drawImages.length);
    expect(reversedRec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
  });

  it('uses structured direct and linked floor/wall paint with direct noFill precedence', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color: '112233' },
        { position: 1, color: 'DDEEFF' },
      ],
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        floor: { fillPaints: [gradient], fillPaintAuthored: true },
        wall: { fillPaints: [gradient], fillPaintAuthored: true },
      },
      threeD: {
        rotationX: 15,
        rotationY: 20,
        floor: { fillHidden: true },
        sideWall: { style: { fillPaints: [gradient], fillPaintAuthored: true } },
      },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    // Direct floor noFill suppresses the linked floor role. The direct side
    // wall and linked back wall each resolve one bounded gradient recipe.
    expect(rec.gradients).toHaveLength(2);
  });

  it('uses plotArea3D for a 3-D plot unless direct plot-area paint wins', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color: '112233' },
        { position: 1, color: 'DDEEFF' },
      ],
    };
    const render = (direct: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A', 'B'],
        chartStyleRoles: {
          plotArea3D: { fillPaints: [gradient], fillPaintAuthored: true },
        },
        plotAreaBg: direct ? 'FF0000' : null,
        plotAreaFillPaintAuthored: direct,
        threeD: { rotationX: 15, rotationY: 20 },
        series: [series({ values: [5, 15] })],
      }), RECT, 1);
      return rec;
    };
    expect(render(false).gradients).toHaveLength(1);
    expect(render(true).gradients).toHaveLength(0);
  });

  it.each(['clusteredBar', 'pie'] as const)(
    'applies linked dataPoint3D paint to %s and keeps direct point noFill authoritative',
    chartType => {
      const gradient = {
        fillType: 'gradient' as const,
        gradType: 'linear' as const,
        angle: 0,
        stops: [
          { position: 0, color: '112233' },
          { position: 1, color: 'DDEEFF' },
        ],
      };
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        valAxisMajorGridlines: false,
        chartStyleRoles: {
          dataPoint3D: {
            fillPaints: [gradient],
            fillPaintAuthored: true,
            linePaints: [gradient],
            linePaintAuthored: true,
            lineWidthEmu: 12_700,
          },
        },
        threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
        series: [series({
          values: [5, 15],
          dataPointOverrides: [{
            idx: 0,
            fillHidden: true,
            lineHidden: true,
            chartexStyle: { fillHidden: true, lineHidden: true },
          }],
        })],
      }), RECT, 1);
      // Only point 1 consumes the linked fill and outline recipes. Each is
      // resolved once for the complete datum, not once for every mesh face.
      expect(rec.gradients).toHaveLength(2);
    },
  );

  it('preflights percent-stack paint without overflowing finite magnitudes', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      chartStyleRoles: {
        dataPoint3D: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096, color: '112233',
            })),
          }],
          fillPaintAuthored: true,
        },
      },
      threeD: { rotationX: 15, rotationY: 20 },
      series: [
        series({ values: [Number.MAX_VALUE] }),
        series({ values: [Number.MAX_VALUE] }),
      ],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it('does not charge saturated stacked-bar segments that paint no mesh', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: Array.from({ length: 4_096 }, (_, index) => ({
        position: index / 4_095, color: '112233',
      })),
    };
    const values = new Array(129).fill(Number.MAX_VALUE) as number[];
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: values.map((_, index) => `C${index}`),
      chartStyleRoles: {
        dataPoint3D: { fillPaints: [gradient], fillPaintAuthored: true },
      },
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values }), series({ values: [...values] })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('(too many data points)');
    expect(rec.gradients).toHaveLength(values.length);
  });

  it('plans percent-stack denominators once per category', () => {
    let valueReads = 0;
    const measuredSeries = Array.from({ length: 200 }, (_, index) => {
      const values = [index + 1];
      Object.defineProperty(values, 0, {
        configurable: true,
        enumerable: true,
        get: () => {
          valueReads++;
          return index + 1;
        },
      });
      return series({ values });
    });
    renderChart(recordingCtx().ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: measuredSeries,
    }), RECT, 1);
    // A per-datum denominator scan is quadratic (well over 80,000 reads for
    // this boundary). The shared category plan keeps all consumers linear.
    expect(valueReads).toBeLessThan(10_000);
  });

  it('preflights stacked-line paint from the cumulative plotted value', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B'],
      valMin: 10,
      valMax: 20,
      chartStyleRoles: {
        dataPoint3D: {
          linePaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096, color: '112233',
            })),
          }],
          linePaintAuthored: true,
        },
      },
      threeD: { rotationX: 15, rotationY: 20 },
      series: [
        series({ values: [6, 6], showMarker: false }),
        series({ values: [6, 6], showMarker: false }),
      ],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it('indexes linked 3-D bar paint by series instead of category', () => {
    const palette = ['AA0000', '00AA00', '0000AA'].map(color => ({
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color },
        { position: 1, color: 'FFFFFF' },
      ],
    }));
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        dataPoint3D: { fillPaints: palette, fillPaintAuthored: true },
      },
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [
        series({ values: [1] }),
        series({ values: [1, 1, 1] }),
      ],
    }), RECT, 1);
    const firstStops = rec.gradients.map(gradient => gradient.stops[0]?.color);
    expect(firstStops.filter(color => color === 'rgba(170,0,0,1)')).toHaveLength(1);
    expect(firstStops.filter(color => color === 'rgba(0,170,0,1)')).toHaveLength(3);
    expect(firstStops).not.toContain('rgba(0,0,170,1)');
  });

  it('indexes linked 3-D paint by point for a single varyColors bar series', () => {
    const palette = ['AA0000', '00AA00', '0000AA'].map(color => ({
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color },
        { position: 1, color: 'FFFFFF' },
      ],
    }));
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      varyColors: true,
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        dataPoint3D: { fillPaints: palette, fillPaintAuthored: true },
      },
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ values: [1, 1, 1] })],
    }), RECT, 1);
    expect(new Set(rec.gradients.map(gradient => gradient.stops[0]?.color))).toEqual(new Set([
      'rgba(170,0,0,1)',
      'rgba(0,170,0,1)',
      'rgba(0,0,170,1)',
    ]));
  });

  it.each([
    ['line', 1],
    ['area', 2],
  ] as const)('resolves linked dataPoint3D paint once for a 3-D %s series', (chartType, count) => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color: '112233' },
        { position: 1, color: 'DDEEFF' },
      ],
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['A', 'B', 'C'],
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        dataPoint3D: {
          fillPaints: [gradient],
          fillPaintAuthored: true,
          linePaints: [gradient],
          linePaintAuthored: true,
        },
      },
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ values: [5, 15, 10], showMarker: false })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(count);
  });

  it.each(['clusteredBar', 'pie'] as const)(
    'keeps direct point structured/unresolved paint above linked dataPoint3D for %s',
    chartType => {
      const linked = {
        fillType: 'gradient' as const, gradType: 'linear' as const, angle: 0,
        stops: [{ position: 0, color: '112233' }, { position: 1, color: 'DDEEFF' }],
      };
      const direct = {
        fillType: 'gradient' as const, gradType: 'linear' as const, angle: 90,
        stops: [{ position: 0, color: 'AA0000' }, { position: 1, color: 'FFCCCC' }],
      };
      const render = (chartexStyle: NonNullable<ChartSeries['dataPointOverrides']>[number]['chartexStyle']) => {
        const rec = recordingCtx();
        renderChart(rec.ctx, baseModel({
          chartType,
          categories: ['A'],
          valAxisMajorGridlines: false,
          chartStyleRoles: {
            dataPoint3D: {
              fillPaints: [linked], fillPaintAuthored: true,
              linePaints: [linked], linePaintAuthored: true,
            },
          },
          threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
          series: [series({
            values: [10],
            dataPointOverrides: [{ idx: 0, chartexStyle }],
          })],
        }), RECT, 1);
        return rec;
      };

      const authored = render({
        fillPaints: [direct], fillPaintAuthored: true,
        linePaints: [direct], linePaintAuthored: true,
      });
      expect(authored.gradients).toHaveLength(2);
      expect(authored.gradients.every(gradient =>
        gradient.stops[0]?.color === 'rgba(170,0,0,1)'
      )).toBe(true);
      const unresolved = render({ fillPaintAuthored: true, linePaintAuthored: true });
      expect(unresolved.gradients).toHaveLength(0);
    },
  );

  it.each([
    ['line', 1],
    ['area', 2],
  ] as const)(
    'keeps direct series structured/unresolved paint above linked dataPoint3D for %s',
    (chartType, expectedGradients) => {
      const linked = {
        fillType: 'gradient' as const, gradType: 'linear' as const, angle: 0,
        stops: [{ position: 0, color: '112233' }, { position: 1, color: 'DDEEFF' }],
      };
      const direct = {
        fillType: 'gradient' as const, gradType: 'linear' as const, angle: 90,
        stops: [{ position: 0, color: 'AA0000' }, { position: 1, color: 'FFCCCC' }],
      };
      const render = (chartexStyle: ChartSeries['chartexStyle']) => {
        const rec = recordingCtx();
        renderChart(rec.ctx, baseModel({
          chartType,
          categories: ['A', 'B'],
          valAxisMajorGridlines: false,
          chartStyleRoles: {
            dataPoint3D: {
              fillPaints: [linked], fillPaintAuthored: true,
              linePaints: [linked], linePaintAuthored: true,
            },
          },
          threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
          series: [series({ values: [5, 10], showMarker: false, chartexStyle })],
        }), RECT, 1);
        return rec;
      };

      const authored = render({
        fillPaints: [direct], fillPaintAuthored: true,
        linePaints: [direct], linePaintAuthored: true,
      });
      expect(authored.gradients).toHaveLength(expectedGradients);
      expect(authored.gradients.every(gradient =>
        gradient.stops[0]?.color === 'rgba(170,0,0,1)'
      )).toBe(true);
      expect(render({ fillPaintAuthored: true, linePaintAuthored: true }).gradients)
        .toHaveLength(0);
    },
  );

  it('rejects oversized 3-D datum paint work before resolving any point recipe', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: Array.from({ length: 4_096 }, (_, index) => ({
        position: index / 4_095,
        color: index % 2 ? '112233' : 'DDEEFF',
      })),
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: Array.from({ length: 257 }, (_, index) => `C${index}`),
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        dataPoint3D: { fillPaints: [gradient], fillPaintAuthored: true },
      },
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ values: new Array(257).fill(1) })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it('rejects an oversized linked 3-D wall recipe before painting the chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      chartStyleRoles: {
        wall: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096, color: '112233',
            })),
          }],
          fillPaintAuthored: true,
        },
      },
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [1] })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toEqual(['(too many data points)']);
  });

  it('does not charge an oversized 3-D line recipe when the series has no geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      chartStyleRoles: {
        dataPoint3D: {
          linePaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096, color: '112233',
            })),
          }],
          linePaintAuthored: true,
        },
      },
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [null, Number.NaN], showMarker: false })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).not.toContain('(too many data points)');
  });

  it.each([
    ['zero bar', 'clusteredBar', [0], null, null],
    ['fully clipped bar', 'clusteredBar', [1], 2, 3],
    ['fully clipped line', 'line', [1, 2], 3, 4],
    ['line collapsed at the lower boundary', 'line', [1, 2], 2, 3],
  ] as const)(
    'does not charge unused structured paint for a %s',
    (_name, chartType, values, valMin, valMax) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: values.map((_, index) => `C${index}`),
        valMin,
        valMax,
        chartStyleRoles: {
          dataPoint3D: {
            fillPaints: [{
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4_097 }, (_, index) => ({
                position: index / 4_096, color: '112233',
              })),
            }],
            fillPaintAuthored: true,
            linePaints: [{
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4_097 }, (_, index) => ({
                position: index / 4_096, color: '445566',
              })),
            }],
            linePaintAuthored: true,
          },
        },
        threeD: { rotationX: 15, rotationY: 20 },
        series: [series({ values: [...values], showMarker: false })],
      }), RECT, 1);
      expect(rec.gradients).toHaveLength(0);
      expect(rec.texts.map(text => text.text)).not.toContain('(too many data points)');
    },
  );

  it('paints authored floor/side/back CT_Surface rules and keeps standard depth substantial', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      valAxisMajorGridlines: false,
      threeD: {
        rotationX: 15, rotationY: 20, heightPercent: 100,
        depthPercent: 100, perspective: 30, barGrouping: 'standard',
        floor: {
          fillHidden: true, lineColor: '000000', lineWidthEmu: 3175,
          lineDash: 'solid', lineHidden: false,
        },
        sideWall: {
          fillHidden: true, lineColor: 'FF0000', lineWidthEmu: 12700,
          lineDash: 'dash', lineHidden: false,
        },
        backWall: {
          fillHidden: true, lineColor: '00FF00', lineWidthEmu: 12700,
          lineDash: 'solid', lineHidden: false,
        },
      },
      series: [
        series({ name: 'One', values: [8, 6, 2, 2] }),
        series({ name: 'Two', values: [7, 5, 3, 1] }),
        series({ name: 'Three', values: [3, 3, 5, 6] }),
      ],
    }), RECT, 1);
    const floor = rec.strokes.filter(stroke => stroke.ss === '#000000');
    const side = rec.strokes.filter(stroke => stroke.ss === '#FF0000');
    const back = rec.strokes.filter(stroke => stroke.ss === '#00FF00');
    expect(floor.length).toBeGreaterThan(0);
    expect(side.length).toBeGreaterThan(0);
    expect(back.length).toBeGreaterThan(0);
    expect(side.every(stroke => stroke.dash.length > 0)).toBe(true);
    const length = (a: { x: number; y: number }, b: { x: number; y: number }) =>
      Math.hypot(a.x - b.x, a.y - b.y);
    const sideEdgeLengths = side.flatMap(stroke =>
      stroke.points.slice(1).map((point, index) => length(stroke.points[index], point))
    ).filter(edgeLength => edgeLength > 1e-6).sort((left, right) => left - right);
    const projectedDepth = sideEdgeLengths[0];
    const projectedHeight = sideEdgeLengths.at(-1) as number;
    // Office's measured default standard-Bar axes are 8.1:8.1:2.6. Keep the
    // final projected depth/height ratio close to 0.321 without applying this
    // compatibility calibration to clustered/stacked bars.
    expect(projectedDepth / projectedHeight).toBeGreaterThan(0.29);
    expect(projectedDepth / projectedHeight).toBeLessThan(0.36);
  });

  it('declines unsupported 3-D families so their canonical renderer remains available', () => {
    const rec = recordingCtx();
    expect(renderSimpleThreeDChart(rec.ctx, baseModel({
      chartType: 'scatter',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [1, 2] })],
    }), RECT, 1)).toBe(false);
  });

  it('does not apply the 3-D mesh budget before an unsupported family falls back', () => {
    const rec = recordingCtx();
    const values = new Array(400).fill(1);
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: values.map((_, index) => String(index)),
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({ values })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('(too many data points)');
  });

  it('places clustered series side by side on the projected category axis', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [
        series({ color: 'FF0000', values: [10] }),
        series({ color: '00FF00', values: [10] }),
      ],
    }), RECT, 1);
    const boundsFor = (color: string) => {
      const points = rec.filledPaths
        .filter(path => isMaterialColor(path.fillStyle, color))
        .flatMap(path => path.points);
      return {
        minX: Math.min(...points.map(point => point.x)),
        maxX: Math.max(...points.map(point => point.x)),
        maxY: Math.max(...points.map(point => point.y)),
      };
    };
    const red = boundsFor('FF0000');
    const green = boundsFor('00FF00');
    // The category axis runs down/right in the default projection. Excel's
    // clustered series follow that edge; the former per-series Z placement
    // instead moved later series up/right and made them overlap.
    expect(green.minX).toBeGreaterThan(red.minX);
    expect(green.maxX).toBeGreaterThan(red.maxX);
    expect(green.maxY).toBeGreaterThan(red.maxY);
  });

  it('places standard 3-D bar series on distinct series-axis depth slots', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
        barGrouping: 'standard',
        seriesAxis: {
          title: 'Depth series', hidden: false, orientation: 'minMax',
          tickLabelPos: 'low', tickLabelSkip: 1, tickMarkSkip: 1,
          majorTickMark: 'out', fontColor: '123456', fontSizeHpt: 800,
          fontBold: false, lineColor: '654321', lineWidthEmu: 12700,
          lineHidden: false, titleFontSizeHpt: 900, titleFontBold: true,
        },
      } as ChartModel['threeD'],
      series: [
        series({ name: 'North', color: 'FF0000', values: [10] }),
        series({ name: 'South', color: '00FF00', values: [10] }),
      ],
    }), RECT, 1);
    const boundsFor = (color: string) => {
      const points = rec.filledPaths
        .filter(path => isMaterialColor(path.fillStyle, color))
        .flatMap(path => path.points);
      return {
        minX: Math.min(...points.map(point => point.x)),
        maxX: Math.max(...points.map(point => point.x)),
        minY: Math.min(...points.map(point => point.y)),
        maxY: Math.max(...points.map(point => point.y)),
      };
    };
    const red = boundsFor('FF0000');
    const green = boundsFor('00FF00');
    // Both series reuse one category footprint. Their projected separation is
    // therefore the camera's Z vector (right/up), not an adjacent X-axis slot
    // (right/down) as in the clustered test above.
    expect(green.minX).toBeGreaterThan(red.minX);
    expect(green.maxX).toBeGreaterThan(red.maxX);
    expect(green.minY).toBeLessThan(red.minY);
    expect(green.maxY).toBeLessThan(red.maxY);
    const standardCenterDistance = Math.hypot(
      (green.minX + green.maxX - red.minX - red.maxX) / 2,
      (green.minY + green.maxY - red.minY - red.maxY) / 2,
    );
    const clustered = recordingCtx();
    renderChart(clustered.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
        barGrouping: 'clustered',
      },
      series: [
        series({ color: 'FF0000', values: [10] }),
        series({ color: '00FF00', values: [10] }),
      ],
    }), RECT, 1);
    const clusteredCenter = (color: string) => {
      const points = clustered.filledPaths
        .filter(path => isMaterialColor(path.fillStyle, color))
        .flatMap(path => path.points);
      return {
        x: (Math.min(...points.map(point => point.x)) + Math.max(...points.map(point => point.x))) / 2,
        y: (Math.min(...points.map(point => point.y)) + Math.max(...points.map(point => point.y))) / 2,
      };
    };
    const clusteredRed = clusteredCenter('FF0000');
    const clusteredGreen = clusteredCenter('00FF00');
    const clusteredCenterDistance = Math.hypot(
      clusteredGreen.x - clusteredRed.x,
      clusteredGreen.y - clusteredRed.y,
    );
    // A standard 3-D group uses a real series axis. Office's authored default
    // leaves that axis visibly deep; it is not the shallow single slab used by
    // a clustered group whose series are already separated on the category
    // axis.
    expect(standardCenterDistance).toBeGreaterThan(clusteredCenterDistance * 0.25);
    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining([
      'North', 'South', 'Depth series',
    ]));
    expect(rec.texts.find(text => text.text === 'North')?.fillStyle).toBe('#123456');
  });

  it('applies the linked series-axis dash to a standard 3-D group', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      chartStyleRoles: {
        seriesAxis: {
          lineColors: ['654321'], lineWidthEmu: 25_400, lineDash: 'dashDot',
          fontSizeHpt: 700, fontBold: true, fontColor: 'AABBCC', fontFace: 'Series Face',
        },
      },
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
        barGrouping: 'standard',
        seriesAxis: {
          hidden: false, orientation: 'minMax', majorTickMark: 'out', lineHidden: false,
        },
      },
      series: [
        series({ name: 'North', color: 'FF0000', values: [10] }),
        series({ name: 'South', color: '00FF00', values: [10] }),
      ],
    }), RECT, 1);
    expect(rec.segs.some(segment =>
      segment.ss === '#654321' && segment.lw === 2 && segment.dash.length === 4
    )).toBe(true);
    const seriesAxisLabel = rec.texts.find(text => text.text === 'North');
    expect(seriesAxisLabel).toMatchObject({ fillStyle: '#AABBCC' });
    expect(seriesAxisLabel?.font).toContain('bold 7px "Series Face"');
  });

  it('depth-sorts crossing 3-D line segments instead of painting whole series atomically', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C', 'D'],
      valMin: 0,
      valMax: 100,
      threeD: { rotationX: 80, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [
        series({ lineColor: 'FF0000', values: [0, 100, 0, 100] }),
        series({ lineColor: '00FF00', values: [100, 0, 100, 0] }),
      ],
    }), RECT, 1);
    const sequence = rec.paintEvents
      .filter((event): event is FillPaintEvent => event.kind === 'fill')
      .map(event => isMaterialColor(event.fillStyle, 'FF0000')
        ? 'red'
        : isMaterialColor(event.fillStyle, '00FF00') ? 'green' : null)
      .filter((color): color is 'red' | 'green' => color != null);
    expect(sequence.length).toBeGreaterThanOrEqual(6);
    expect(sequence.some((color, index) => index > 0
      && index + 1 < sequence.length
      && color !== sequence[index - 1]
      && color === sequence[index + 1])).toBe(true);
  });

  it('extrudes each 3-D line stroke through its series depth interval', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 40,
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100,
        gapDepthPercent: 150, perspective: 30,
      },
      series: [series({ lineColor: 'FF0000', lineWidthEmu: 38_100, values: [20, 30] })],
    }), RECT, 1);
    const redFaces = materialFills(rec, 'FF0000');
    // A flat stroke emits one screen-space polygon. Excel renders Line3D as a
    // solid ribbon, so its camera projection exposes multiple shaded faces.
    expect(redFaces.length).toBeGreaterThan(1);
    expect(new Set(redFaces.map(event => event.fillStyle)).size).toBeGreaterThan(1);
  });

  it('applies a direct 3-D line point style to the segment ending at that point', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 10,
      threeD: { rotationX: 0, rotationY: 0, perspective: 0 },
      series: [series({
        values: [2, 8, 4],
        lineColor: '0000FF',
        showMarker: false,
        dataPointOverrides: [{ idx: 1, lineColor: 'FF0000' }],
      })],
    }), RECT, 1);
    const red = rec.filledPaths.filter(path => path.fillStyle === '#FF0000');
    const blue = rec.filledPaths.filter(path => path.fillStyle === '#0000FF');
    expect(red.length).toBeGreaterThan(0);
    expect(blue.length).toBeGreaterThan(0);
    const centerX = (paths: typeof red) => {
      const points = paths.flatMap(path => path.points);
      return points.reduce((sum, point) => sum + point.x, 0) / points.length;
    };
    // Office assigns dPt idx=1 to the incoming A→B segment. The following
    // B→C segment remains owned by the series style.
    expect(centerX(red)).toBeLessThan(centerX(blue));
  });

  it('suppresses only the incoming 3-D line segment for direct point noFill', () => {
    const render = (lineHidden: boolean | undefined) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line', categories: ['A', 'B', 'C'], valMin: 0, valMax: 10,
        threeD: { rotationX: 0, rotationY: 0, perspective: 0 },
        series: [series({
          values: [2, 8, 4], lineColor: '0000FF', showMarker: false,
          dataPointOverrides: lineHidden == null ? undefined : [{ idx: 1, lineHidden }],
        })],
      }), RECT, 1);
      return rec.filledPaths.filter(path => path.fillStyle === '#0000FF');
    };
    const baseline = render(undefined);
    const hidden = render(true);
    expect(hidden.length).toBeGreaterThan(0);
    expect(hidden.length).toBeLessThan(baseline.length);
    const centerX = (paths: typeof hidden) => paths
      .flatMap(path => path.points)
      .reduce((sum, point) => sum + point.x, 0)
      / paths.flatMap(path => path.points).length;
    // idx=1 hides A→B. The later B→C segment remains visible.
    expect(centerX(hidden)).toBeGreaterThan(centerX(baseline));
  });

  it('resolves one structured point paint for its incoming 3-D line segment', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color: 'FF0000' },
        { position: 1, color: 'FFFFFF' },
      ],
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [2, 8, 4], lineColor: '0000FF', showMarker: false,
        dataPointOverrides: [{
          idx: 1,
          chartexStyle: { linePaints: [gradient], linePaintAuthored: true },
        }],
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(1);
    expect(rec.texts.map(text => text.text)).not.toContain('(too many data points)');
  });

  it('rejects an oversized direct point line paint before any 3-D geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [2, 8], lineColor: '0000FF', showMarker: false,
        dataPointOverrides: [{
          idx: 1,
          chartexStyle: {
            linePaints: [{
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4_097 }, (_, index) => ({
                position: index / 4_096, color: 'FF0000',
              })),
            }],
            linePaintAuthored: true,
          },
        }],
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toEqual(['(too many data points)']);
  });

  it('does not apply direct point paint to a 3-D area body', () => {
    const gradient = {
      fillType: 'gradient' as const, gradType: 'linear' as const, angle: 0,
      stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: 'FFFFFF' }],
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area', categories: ['A', 'B', 'C'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [2, 8, 4], color: '4472C4', showMarker: false,
        dataPointOverrides: [{
          idx: 1,
          chartexStyle: {
            fillPaints: [gradient], fillPaintAuthored: true,
            linePaints: [gradient], linePaintAuthored: true,
          },
        }],
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
  });

  it('builds stacked 3-D area layers from cumulative boundaries', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [
        series({ lineColor: 'FF0000', values: [20, 20] }),
        series({ lineColor: '00FF00', values: [10, 10] }),
      ],
    }), RECT, 1);
    const first = rec.filledPaths.filter(path => path.fillStyle === '#FF0000');
    const second = rec.filledPaths.filter(path => path.fillStyle === '#00FF00');
    expect(first.length).toBeGreaterThan(0);
    expect(second.length).toBeGreaterThan(0);
    // The second source value is smaller (10 vs 20), so a raw-series painter
    // would put it lower. Correct stacking paints its upper boundary at 30.
    expect(Math.max(...second.flatMap(path => path.points.map(point => point.y))))
      .toBeLessThan(Math.min(...first.flatMap(path => path.points.map(point => point.y))));
    const fillRec = recordingCtx();
    renderChart(fillRec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [
        series({ color: 'FF0000', values: [20, 20] }),
        series({ color: '00FF00', values: [10, 10] }),
      ],
    }), RECT, 1);
    expect(materialFills(fillRec, 'FF0000').length).toBeGreaterThan(0);
    expect(materialFills(fillRec, '00FF00').length).toBeGreaterThan(0);
  });

  it('extrudes each 3-D area series into camera-projected depth faces', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 40,
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100,
        gapDepthPercent: 150, perspective: 30,
      },
      series: [series({ color: 'FF0000', lineHidden: true, values: [20, 30] })],
    }), RECT, 1);
    const redFaces = materialFills(rec, 'FF0000');
    // A flat area emits one quadrilateral for this two-category run.  The
    // Office model is a solid ribbon: at least the broad face and one of its
    // projected ridge/end faces must be visible and depth-shaded.
    expect(redFaces.length).toBeGreaterThan(1);
    expect(new Set(redFaces.map(event => String(event.fillStyle))).size).toBeGreaterThan(1);
  });

  it('uses shaded extrusion faces instead of inventing an automatic 3-D area outline', () => {
    const automatic = strokedPolylineCtx();
    renderChart(automatic.ctx, baseModel({
      chartType: 'area', categories: ['A', 'B'], valMin: 0, valMax: 40,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ color: '4472C4', values: [20, 30] })],
    }), RECT, 1);
    expect(automatic.strokes.some(stroke => stroke.ss === '#2f5089')).toBe(false);

    const authored = recordingCtx();
    renderChart(authored.ctx, baseModel({
      chartType: 'area', categories: ['A', 'B'], valMin: 0, valMax: 40,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ color: '4472C4', lineColor: 'FF0000', values: [20, 30] })],
    }), RECT, 1);
    expect(authored.filledPaths.some(path => path.fillStyle === '#FF0000')).toBe(true);
  });

  it.each([
    ['reversed categories', { catAxisOrientation: 'maxMin' as const }, [20, 30]],
    ['negative values', {}, [-20, -30]],
    ['zero crossing', {}, [-20, 30]],
  ])('keeps extruded 3-D area faces outward for %s', (_name, extra, values) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area', categories: ['A', 'B'],
      valMin: -40, valMax: 40,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ color: 'FF0000', lineHidden: true, values })],
      ...extra,
    }), RECT, 1);
    expect(materialFills(rec, 'FF0000').length).toBeGreaterThan(1);
  });

  it('keeps positive-domain and logarithmic 3-D area geometry inside the plot', () => {
    for (const model of [
      baseModel({
        chartType: 'area', categories: ['A', 'B'],
        valMin: 50, valMax: 100, valAxisMajorUnit: 50,
        threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
        series: [series({ values: [50, 100] })],
      }),
      baseModel({
        chartType: 'area', categories: ['A', 'B'],
        valAxisLogBase: 10,
        threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
        series: [series({ values: [10, 1000] })],
      }),
    ]) {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, model, RECT, 1);
      const points = rec.strokes.flatMap(stroke => stroke.points);
      expect(points.length).toBeGreaterThan(0);
      expect(points.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))).toBe(true);
      expect(points.every(point =>
        point.x >= RECT.x - 1 && point.x <= RECT.x + RECT.w + 1
        && point.y >= RECT.y - 1 && point.y <= RECT.y + RECT.h + 1
      )).toBe(true);
    }
  });

  it('starts an authored 3-D area ridge where the surface enters explicit bounds', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area', categories: ['Outside', 'Inside'],
      valMin: 0, valMax: 10,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ values: [-10, 10], lineColor: 'FF0000' })],
    }), RECT, 1);
    const ridgePoints = rec.filledPaths
      .filter(path => path.fillStyle === '#FF0000')
      .flatMap(path => path.points);
    expect(ridgePoints.length).toBeGreaterThan(0);
    expect(Math.min(...ridgePoints.map(point => point.x)))
      .toBeGreaterThan(RECT.x + RECT.w * 0.25);
  });

  it('clips 3-D line segments at authored bounds and omits outside markers and labels', () => {
    const chart = baseModel({
      chartType: 'line', categories: ['Low', 'Middle', 'High'],
      valMin: 0, valMax: 10, valAxisMajorUnit: 5,
      showDataLabels: true,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        lineColor: 'FF0000', showMarker: true, markerSymbol: 'circle',
        values: [-10, 5, 20],
      })],
    });
    const rec = recordingCtx();
    renderChart(rec.ctx, chart, RECT, 1);
    const lineSegments = rec.filledPaths.filter(path => isMaterialColor(path.fillStyle, 'FF0000'));
    expect(lineSegments.length).toBeGreaterThanOrEqual(2);
    const labels = recordingCtx();
    renderChart(labels.ctx, chart, RECT, 1);
    expect(labels.texts.some(text => text.text === '5')).toBe(true);
    expect(labels.texts.some(text => text.text === '-10' || text.text === '20')).toBe(false);
  });

  it('lets point delete=false restore one 3-D label from a deleted series collection', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      valMin: 0, valMax: 100, valAxisMajorUnit: 20,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        values: [37, 43],
        seriesDataLabels: {
          deleted: true,
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
        },
        dataLabelOverrides: [{ idx: 1, text: '', deleted: false, showVal: true }],
      })],
    }), RECT, 1);

    const texts = rec.texts.map(text => text.text);
    expect(texts).toContain('43');
    expect(texts).not.toContain('37');
  });

  it('insets 3-D line category labels when crossBetween is between', () => {
    const rec = recordingCtx(() => 5);
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      catAxisLabelRotation: 0,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [10, 20, 30] })],
    }), RECT, 1);
    const category = rec.texts.find(text => text.text === 'A');
    const zero = rec.texts.find(text => text.text === '0');
    expect(category).toBeDefined();
    expect(zero).toBeDefined();
    // ECMA-376 `crossBetween=between` reserves half a category step at each
    // end. The first label/data plane must therefore sit inside the value-axis
    // edge instead of collapsing onto it.
    expect((category?.x ?? 0) - (zero?.x ?? 0)).toBeGreaterThan(20);
  });

  it('projects authored 3-D axis crossings before drawing rules, ticks and labels', () => {
    const render = (
      catCross: number | 'min' | 'max',
      valCross: 'min' | 'max',
      valOrientation: 'minMax' | 'maxMin' = 'minMax',
    ) => {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A', 'B', 'C'],
        valMin: -10,
        valMax: 10,
        valAxisOrientation: valOrientation,
        catAxisCrossesAt: typeof catCross === 'number' ? catCross : null,
        catAxisCrosses: typeof catCross === 'string' ? catCross : null,
        valAxisCrosses: valCross,
        catAxisLineColor: 'FF0000',
        valAxisLineColor: '00FF00',
        catAxisMajorTickMark: 'none',
        valAxisMajorTickMark: 'none',
        valAxisMajorGridlines: false,
        threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
        series: [series({ values: [-5, 0, 5] })],
      }), RECT, 1);
      const rule = (color: string) => rec.strokes
        .filter(stroke => stroke.ss === color && stroke.points.length === 2)
        .reduce((longest, stroke) => {
          const length = (item: typeof stroke) => Math.hypot(
            item.points[1].x - item.points[0].x,
            item.points[1].y - item.points[0].y,
          );
          return !longest || length(stroke) > length(longest) ? stroke : longest;
        }, null as (typeof rec.strokes)[number] | null);
      return { category: rule('#FF0000'), value: rule('#00FF00') };
    };
    const minimum = render(-10, 'min');
    const maximum = render(10, 'max');
    const reversedMinimum = render(-10, 'min', 'maxMin');
    const authoredReversedMinimum = render('min', 'min', 'maxMin');
    expect(minimum.category).not.toBeNull();
    expect(maximum.category).not.toBeNull();
    expect(Math.abs(
      ((minimum.category?.points[0].y ?? 0) + (minimum.category?.points[1].y ?? 0)) / 2
      - ((maximum.category?.points[0].y ?? 0) + (maximum.category?.points[1].y ?? 0)) / 2
    )).toBeGreaterThan(20);
    const meanY = (stroke: typeof minimum.category) =>
      ((stroke?.points[0].y ?? 0) + (stroke?.points[1].y ?? 0)) / 2;
    expect(meanY(reversedMinimum.category)).toBeCloseTo(meanY(maximum.category), 6);
    expect(meanY(authoredReversedMinimum.category)).toBeCloseTo(meanY(maximum.category), 6);
    expect(Math.abs(
      ((minimum.value?.points[0].x ?? 0) + (minimum.value?.points[1].x ?? 0)) / 2
      - ((maximum.value?.points[0].x ?? 0) + (maximum.value?.points[1].x ?? 0)) / 2
    )).toBeGreaterThan(20);
  });

  it.each([20, 180, 200])('keeps low/high 3-D value labels on screen-left/right at rotY=%s', rotationY => {
    const labelX = (position: 'low' | 'high') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line', categories: ['A', 'B'],
        valMin: 0, valMax: 20, valAxisMajorUnit: 10,
        catAxisCrossesAt: 10, valAxisTickLabelPos: position,
        threeD: { rotationX: 15, rotationY },
        series: [series({ values: [5, 15] })],
      }), RECT, 1);
      return rec.texts.find(text => text.text === '0')?.x;
    };
    expect(labelX('low')).toBeLessThan(labelX('high') as number);
  });

  it('clamps extreme finite 3-D category-axis crossings before projection', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      valAxisCrossesAt: Number.MAX_VALUE,
      valAxisLineColor: 'FF0000',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [5, 15] })],
    }), RECT, 1);
    const vertices = rec.strokes.flatMap(stroke => stroke.points);
    expect(vertices.length).toBeGreaterThan(0);
    expect(vertices.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))).toBe(true);
  });

  it('automatically angles dense 3-D category labels but honors explicit zero', () => {
    const chart = baseModel({
      chartType: 'area',
      categories: Array.from({ length: 8 }, (_, index) => `1900-01-${String(index + 1).padStart(2, '0')}`),
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: [10, 12, 16, 15, 21, 24, 28, 30] })],
    });
    const automatic = recordingCtx();
    renderChart(automatic.ctx, chart, RECT, 1);
    expect(automatic.rotations).toContain(-Math.PI / 4);

    const authored = recordingCtx();
    renderChart(authored.ctx, { ...chart, catAxisLabelRotation: 0 }, RECT, 1);
    expect(authored.rotations).toHaveLength(0);
  });

  it('reverses 3-D category geometry and labels for maxMin orientation', () => {
    const render = (reversed: boolean) => {
      const rec = recordingCtx(() => 5);
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['First', 'Last'],
        catAxisLabelRotation: 0,
        catAxisOrientation: reversed ? 'maxMin' : null,
        threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
        series: [series({ values: [10, 30] })],
      }), RECT, 1);
      return rec;
    };
    const ordinary = render(false);
    const reversed = render(true);
    const textX = (rec: ReturnType<typeof render>, label: string) =>
      rec.texts.find(text => text.text === label)?.x ?? Number.NaN;
    expect(textX(ordinary, 'First')).toBeLessThan(textX(ordinary, 'Last'));
    expect(textX(reversed, 'First')).toBeGreaterThan(textX(reversed, 'Last'));
  });

  it('keeps all-negative 3-D percent stacks on the negative side of zero', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedAreaPct',
      categories: ['A', 'B'],
      valAxisFormatCode: '0%',
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [
        series({ values: [-20, -10] }),
        series({ values: [-30, -40] }),
      ],
    }), RECT, 1);
    const axisLabels = rec.texts
      .map(text => text.text)
      .filter(text => /^-?\d+%$/.test(text));
    expect(axisLabels).toContain('-100%');
    expect(axisLabels).toContain('0%');
    expect(axisLabels.some(text => !text.startsWith('-') && text !== '0%')).toBe(false);
    expect(rec.texts.every(text => Number.isFinite(text.x) && Number.isFinite(text.y))).toBe(true);
  });

  it('keeps non-finite and extreme 3-D values out of Canvas coordinates', () => {
    for (const chartType of ['clusteredBar', 'line', 'stackedArea', 'stackedAreaPct'] as const) {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        threeD: { rotationX: 15, rotationY: 20, depthPercent: 2000, perspective: 240 },
        series: [
          series({ values: [Number.MAX_VALUE, Number.NaN, Number.POSITIVE_INFINITY] }),
          series({ values: [Number.MAX_VALUE, -Number.MAX_VALUE, 1] }),
        ],
      }), RECT, 1);
      expect(rec.strokes.flatMap(stroke => stroke.points).every(point =>
        Number.isFinite(point.x) && Number.isFinite(point.y)
      ), chartType).toBe(true);
    }
  });

  it('rejects oversized 3-D input before expanding projected faces', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: Array.from({ length: 10_001 }, (_, index) => String(index)),
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      series: [series({ values: new Array(10_001).fill(1) })],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('(too many data points)');
    expect(rec.paintEvents.some(event => event.kind === 'fill')).toBe(false);
  });

  it.each([
    'box', 'cylinder', 'cone', 'coneToMax', 'pyramid', 'pyramidToMax',
  ])('projects authored 3-D bar shape %s through the shared camera', shape => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      threeD: {
        rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30, shape,
      },
      series: [series({ values: [20] })],
    }), RECT, 1);
    const coloredFaces = materialFills(rec, '4472C4');
    expect(coloredFaces.length).toBeGreaterThan(0);
  });

  it.each(['cone', 'pyramid', 'coneToMax', 'pyramidToMax'])(
    'clips %s at the original solid cross-section instead of regenerating it',
    shape => {
      const renderedWidth = (valMin: number) => {
        const rec = recordingCtx();
        renderChart(rec.ctx, baseModel({
          chartType: 'clusteredBar', categories: ['A'], valMin, valMax: 100,
          threeD: { rotationX: 15, rotationY: 20, shape },
          series: [series({ color: '4472C4', values: [75] })],
        }), RECT, 1);
        const points = rec.filledPaths
          .filter(path => isMaterialColor(path.fillStyle, '4472C4'))
          .flatMap(path => path.points);
        return Math.max(...points.map(point => point.x)) - Math.min(...points.map(point => point.x));
      };
      const full = renderedWidth(0);
      const clipped = renderedWidth(50);
      expect(full).toBeGreaterThan(0);
      expect(clipped / full).toBeLessThan(0.72);
      expect(clipped / full).toBeGreaterThan(0.18);
    },
  );

  it('charges data-label legend keys to the chart-wide marker paint budget', () => {
    const rec = recordingCtx();
    const count = 256;
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, (_, index) => index + 1),
        showMarker: true, markerSymbol: 'circle',
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: Array.from({ length: 2049 }, (_, index) => ({
            position: index / 2048,
            color: '112233',
          })),
        },
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          showLegendKey: true,
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
  });

  it('refuses repeated data-label shape gradients before partial callout paint', () => {
    const rec = recordingCtx();
    const count = 257;
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, () => 1),
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          labelBox: {
            fillPaint: {
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4096 }, (_, index) => ({
                position: index / 4095,
                color: '112233',
              })),
            },
          },
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
  });

  it('bounds ordinary 2-D label shape work at the 256/257 gradient boundary', () => {
    const build = (count: number): ChartModel => baseModel({
      chartType: 'line',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, (_, index) => index + 1),
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          labelBox: {
            fillPaint: {
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4096 }, (_, index) => ({
                position: index / 4095,
                color: '112233',
              })),
            },
          },
        },
      })],
    });
    expect(chartLabelPaintWorkCount(build(256), undefined)).toBe(1_048_576);
    expect(chartLabelPaintWorkCount(build(257), undefined)).toBe(1_048_577);
    const rec = recordingCtx();
    renderChart(rec.ctx, build(257), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
  });

  it('paints the label shape but not glyphs for direct text noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        values: [42],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          fontPaintAuthored: true,
          fontHidden: true,
          labelBox: { fill: 'D9EAF7' },
        },
      })],
    }), RECT, 1);
    expect(rec.rects.some(rect => rect.fs === '#D9EAF7')).toBe(true);
    expect(rec.texts.some(text => text.text === '42')).toBe(false);
  });

  it('rejects oversized ChartEx public structure before linked-style projection', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'boxWhisker',
      chartStyleRoles: { dataLabel: { fontColor: '112233' } },
      series: Array.from({ length: 10_001 }, (_, index) => series({
        name: String(index), values: [],
      })),
      chartexBox: { categories: [], series: [] },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
  });

  it('charges only non-deleted callout labels to the shape-paint budget', () => {
    const rec = recordingCtx();
    const count = 257;
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, () => 1),
        dataLabelOverrides: Array.from({ length: count - 1 }, (_, index) => ({
          idx: index + 1,
          text: '',
          deleted: true,
        })),
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          labelBox: {
            fillPaint: {
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: Array.from({ length: 4096 }, (_, index) => ({
                position: index / 4095,
                color: '112233',
              })),
            },
          },
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(false);
    expect(rec.gradients).toHaveLength(1);
  });

  it('charges linked trendline-label shape paint before any series is painted', () => {
    const rec = recordingCtx();
    const templateSeries = series({
      values: [1, 2], trendLines: [{ trendlineType: 'linear', dispEq: true }],
    });
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: Array.from({ length: 257 }, (_, index) => ({
        ...templateSeries,
        name: `S${index}`,
      })),
      chartStyleRoles: {
        trendlineLabel: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4096 }, (_, index) => ({
              position: index / 4095,
              color: '112233',
            })),
          }],
        },
      },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
  });

  it('does not charge a marker key for an inapplicable scatter data table', () => {
    const count = 256;
    const markerFillPaint = {
      fillType: 'gradient' as const, gradType: 'linear' as const, angle: 0,
      stops: Array.from({ length: 4096 }, (_, index) => ({
        position: index / 4095,
        color: '112233',
      })),
    };
    const model = baseModel({
      chartType: 'scatter',
      categories: Array.from({ length: count }, (_, index) => String(index + 1)),
      dataTable: {
        showHorizontalBorder: false,
        showVerticalBorder: false,
        showOutline: false,
        showKeys: true,
      },
      series: [series({
        values: Array.from({ length: count }, (_, index) => index + 1),
        showMarker: true, markerSymbol: 'circle', markerFillPaint,
      })],
    });
    expect(classicMarkerPaintWorkCount(model)).toBe(1_048_576);
  });

  it('charges direct bubble shape fill and outline recipes once per visible point', () => {
    const model = baseModel({
      chartType: 'bubble', categories: ['0'],
      series: [series({
        values: [0.5], bubbleSizes: [100],
        chartexStyle: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 0.5, color: '778899' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          fillPaintAuthored: true,
          linePaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '000000' },
              { position: 1, color: 'FFFFFF' },
            ],
          }],
          linePaintAuthored: true,
        },
      })],
      catAxisMin: 0, catAxisMax: 1, valMin: 0, valMax: 1,
    });
    expect(classicMarkerPaintWorkCount(model, undefined, 1, RECT)).toBe(5);
  });

  it('charges every bounded bubble3D material stop for plot, legend, and label keys', () => {
    const model = baseModel({
      chartType: 'bubble', showLegend: true, categories: ['0'],
      series: [series({
        values: [0.5], bubbleSizes: [100], bubble3D: true,
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false,
          showPercent: false, showLegendKey: true,
        },
      })],
      catAxisMin: 0, catAxisMax: 1, valMin: 0, valMax: 1,
    });
    // Three gradients consume 4 + 5 + 6 components per painted bubble.
    expect(classicMarkerPaintWorkCount(model, undefined, 1, RECT)).toBe(45);

    model.series[0].chartexStyle = {
      fillHidden: true, fillPaintAuthored: true,
    };
    expect(classicMarkerPaintWorkCount(model, undefined, 1, RECT)).toBe(0);
  });

  it('does not charge structured bubble paint for invisible bubble sizes', () => {
    const count = 257;
    const stops = Array.from({ length: 4_096 }, (_, index) => ({
      position: index / 4_095,
      color: '112233',
    }));
    const model = baseModel({
      chartType: 'bubble',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array<number | null>(count).fill(1),
        bubbleSizes: Array<number | null>(count).fill(0),
        chartexStyle: {
          fillPaints: [{ fillType: 'gradient', gradType: 'linear', angle: 0, stops }],
          fillPaintAuthored: true,
        },
      })],
      catAxisMin: 0, catAxisMax: count, valMin: 0, valMax: 2,
    });
    expect(classicMarkerPaintWorkCount(model, undefined, 1, RECT)).toBe(0);
  });

  it('charges the normal legend marker after plot marker work', () => {
    const count = 256;
    const model = baseModel({
      chartType: 'line', showLegend: true,
      categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, (_, index) => index + 1),
        showMarker: true, markerSymbol: 'circle',
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: Array.from({ length: 4096 }, (_, index) => ({
            position: index / 4095,
            color: '112233',
          })),
        },
      })],
    });
    expect(classicMarkerPaintWorkCount(model)).toBeGreaterThan(1_048_576);
  });

  it('charges exact tiled-image draw repetitions before marker paint', () => {
    const picture = {
      fillType: 'image' as const, stretch: false,
      imagePath: 'xl/media/tile.png', mimeType: 'image/png', dpi: 96,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    const model = baseModel({
      chartType: 'line',
      categories: Array.from({ length: 273 }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: 273 }, (_, index) => index + 1),
        showMarker: true, markerSymbol: 'picture', markerSize: 45,
        markerFillPaint: picture,
      })],
    });
    const bitmap = { width: 1, height: 1 } as unknown as CanvasImageSource;
    expect(chartImageFillPaintWork(picture, () => bitmap, 45, 45, 1)).toBe(3_844);
    expect(classicMarkerPaintWorkCount(model, () => bitmap, 1, RECT))
      .toBeGreaterThan(1_048_576);
  });

  it('does not zero-charge smaller tiled label keys when a large marker size exceeds its cap', () => {
    const count = 5_000;
    const picture = {
      fillType: 'image' as const, stretch: false,
      imagePath: 'xl/media/key-tile.png', mimeType: 'image/png', dpi: 96,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    const model = baseModel({
      chartType: 'line', categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, () => 1),
        showMarker: true, markerSymbol: 'picture', markerSize: 72,
        markerFillPaint: picture,
        dataPointOverrides: Array.from({ length: count }, (_, idx) => ({
          idx, markerSymbol: 'none',
        })),
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false,
          showPercent: false, showLegendKey: true,
        },
      })],
    });
    const bitmap = { width: 1, height: 1 } as unknown as CanvasImageSource;
    expect(classicMarkerPaintWorkCount(model, () => bitmap, 4 / 3, RECT))
      .toBeGreaterThan(1_048_576);
  });

  it('does not fabricate a cap for zero-height 3-D shapes and folds unknown shapes to box', () => {
    const coloredFaceCount = (shape: string, value: number) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar', categories: ['A'],
        threeD: { rotationX: 15, rotationY: 20, shape },
        series: [series({ color: 'FF0000', values: [value] })],
      }), RECT, 1);
      return materialFills(rec, 'FF0000').length;
    };
    expect(coloredFaceCount('cylinder', 0)).toBe(0);
    expect(coloredFaceCount('schema-invalid', 10)).toBe(coloredFaceCount('box', 10));
  });

  it('renders revolved 3-D shapes as a camera-culled, smoothly shaded ring mesh', () => {
    const shadeSet = (shape: 'cylinder' | 'cone') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar', categories: ['A'],
        threeD: { rotationX: 15, rotationY: 20, shape },
        series: [series({ color: '4472C4', values: [20] })],
      }), RECT, 1);
      return new Set(materialFills(rec, '4472C4').map(event => event.fillStyle));
    };

    // A single flat shade makes the ring read as a box/triangle. Office's
    // automatic material has a bright band and progressively darker shoulders;
    // the shared normal-light rule must therefore emit several facet shades.
    expect(shadeSet('cylinder').size).toBeGreaterThanOrEqual(5);
    expect(shadeSet('cone').size).toBeGreaterThanOrEqual(5);
  });

  it('keeps automatic mesh lighting near the authored color while retaining side shading', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({ color: '4472C4', values: [20] })],
    }), RECT, 1);
    const factors = materialFills(rec, '4472C4').map(event => {
      const color = /^#([0-9a-f]{6})$/i.exec(event.fillStyle)?.[1] ?? '000000';
      return parseInt(color.slice(2, 4), 16) / 0x72;
    });
    // Office's automatic material keeps the illuminated band at the authored
    // accent and bottoms the visible shoulder near 78%. A 62% ambient floor
    // made the same real mesh read noticeably muddy.
    expect(Math.max(...factors)).toBeGreaterThanOrEqual(0.98);
    expect(Math.min(...factors)).toBeGreaterThanOrEqual(0.75);
    expect(Math.max(...factors) - Math.min(...factors)).toBeGreaterThan(0.08);
  });

  it('uses the complete gapWidth group for a single 3-D series', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A', 'B', 'C', 'D'],
      barGapWidth: 150,
      threeD: { rotationX: 15, rotationY: 20, gapDepthPercent: 150, shape: 'cylinder' },
      series: [series({ color: '4472C4', values: [10, 10, 10, 10] })],
    }), RECT, 1);
    const faces = rec.filledPaths.filter(path => isMaterialColor(path.fillStyle, '4472C4'));
    const bounds = faces.map(path => ({
      minX: Math.min(...path.points.map(point => point.x)),
      maxX: Math.max(...path.points.map(point => point.x)),
    }));
    const totalMin = Math.min(...bounds.map(bound => bound.minX));
    const totalMax = Math.max(...bounds.map(bound => bound.maxX));
    const categoryPitch = (totalMax - totalMin) / 3;
    const widestFace = Math.max(...bounds.map(bound => bound.maxX - bound.minX));
    // gapWidth=150 leaves a 40% marker group. Projection adds a small visible
    // depth component, so the widest face is close to (but not below) that
    // category-axis fraction. Multi-series clusters split this group instead
    // of shrinking the single-series marker heuristically.
    expect(widestFace / categoryPitch).toBeGreaterThan(0.34);
    expect(widestFace / categoryPitch).toBeLessThan(0.48);
  });

  it('lets a series-level 3-D shape override the chart-group shape', () => {
    const renderFaceCount = (threeDShape: string | null) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar',
        categories: ['A'],
        threeD: {
          rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30,
          shape: 'cylinder',
        },
        series: [series({ threeDShape, values: [20] })],
      }), RECT, 1);
      return materialFills(rec, '4472C4').length;
    };
    expect(renderFaceCount('pyramid')).not.toBe(renderFaceCount(null));
  });

  it('preserves authored 3-D series outline/noFill instead of restoring defaults', () => {
    const outlined = recordingCtx();
    renderChart(outlined.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'box' },
      series: [series({ values: [10], lineColor: 'FF0000', lineWidthEmu: 25400 })],
    }), RECT, 1);
    expect(outlined.filledPaths.some(path => path.fillStyle === '#FF0000')).toBe(true);

    const hidden = recordingCtx();
    renderChart(hidden.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [10, 20], lineColor: 'FF0000', lineHidden: true })],
    }), RECT, 1);
    expect(hidden.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#FF0000'
    )).toBe(false);

    const transparent = recordingCtx();
    renderChart(transparent.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'box' },
      series: [series({ color: '00000000', values: [10], lineHidden: true })],
    }), RECT, 1);
    expect(transparent.paintEvents.some(event =>
      event.kind === 'fill' && String(event.fillStyle).startsWith('rgba(0,0,0,')
    )).toBe(false);
  });

  it('derives an authored cylinder rim from the visible mesh boundary', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({
        values: [20], color: '4472C4', lineColor: 'FF0000', lineWidthEmu: 12_700,
      })],
    }), RECT, 1);
    const authored = rec.filledPaths.filter(path => path.fillStyle === '#FF0000');
    expect(authored.length).toBeGreaterThan(0);
    // Authored rims are expanded into camera-sortable stroke polygons rather
    // than one Canvas stroke per facet. The shared mesh seams therefore never
    // reappear as duplicated Canvas outlines.
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#FF0000'
    )).toBe(false);
    expect(authored.every(path => path.points.length >= 3)).toBe(true);
  });

  it('does not invent a black outline around unstyled 3-D stacked-column faces', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'box' },
      series: [
        series({ color: '156082', values: [10, 12] }),
        series({ color: 'E97132', values: [8, 11] }),
        series({ color: '196B24', values: [12, 14] }),
      ],
    }), RECT, 1);

    // The workbook-level chart border is not a datum outline. With no local
    // <c:ser>/<c:dPt> line authoring, Office distinguishes the stacked slabs
    // only by their fill and shaded faces.
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === 'rgba(0,0,0,0.42)'
    )).toBe(false);
  });

  it('removes the duplicated cap at a shared stacked-cylinder ring', () => {
    const renderBlueFaces = (stacked: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: stacked ? 'stackedBar' : 'clusteredBar',
        categories: ['A'], valMin: 0, valMax: 20,
        threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
        series: stacked
          ? [series({ color: '4472C4', values: [10] }), series({ color: '4472C4', values: [10] })]
          : [series({ color: '4472C4', values: [10] })],
      }), RECT, 1);
      return materialFills(rec, '4472C4').length;
    };
    const oneClosedCylinder = renderBlueFaces(false);
    // Two closed solids would paint exactly twice as many visible faces. The
    // shared end/base plane is one camera-visible duplicate and must disappear.
    expect(renderBlueFaces(true)).toBe(oneClosedCylinder * 2 - 1);
  });

  it('removes the shared zero-plane caps from an opaque signed stack', () => {
    const renderBlueFaces = (values: number[], stacked: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: stacked ? 'stackedBar' : 'clusteredBar',
        categories: ['A'], valMin: -20, valMax: 20,
        threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
        series: values.map(value => series({ color: '4472C4', values: [value] })),
      }), RECT, 1);
      return materialFills(rec, '4472C4').length;
    };
    const positive = renderBlueFaces([10], false);
    const negative = renderBlueFaces([-10], false);
    // Exactly one of the two coincident zero-plane caps is camera-facing in
    // the separate solids. Their signed stack is one continuous mesh there.
    expect(renderBlueFaces([10, -10], true)).toBe(positive + negative - 1);
  });

  it('keeps an opaque stack cap exposed beside a noFill segment', () => {
    const renderNegativeFaces = (withTransparentPositive: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: withTransparentPositive ? 'stackedBar' : 'clusteredBar',
        categories: ['A'], valMin: -20, valMax: 20,
        threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
        series: withTransparentPositive
          ? [
              series({ color: '00000000', lineHidden: true, values: [10] }),
              series({ color: '4472C4', values: [-10] }),
            ]
          : [series({ color: '4472C4', values: [-10] })],
      }), RECT, 1);
      return materialFills(rec, '4472C4').length;
    };
    expect(renderNegativeFaces(true)).toBe(renderNegativeFaces(false));
  });

  it('uses logarithmic model distance for clipped cone cross-sections', () => {
    const renderedWidth = (shape: 'cylinder' | 'cone') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar', categories: ['A'],
        valMin: 1, valMax: 100, valAxisLogBase: 10,
        threeD: { rotationX: 15, rotationY: 20, shape },
        series: [series({ color: '4472C4', values: [10] })],
      }), RECT, 1);
      const points = rec.filledPaths
        .filter(path => isMaterialColor(path.fillStyle, '4472C4'))
        .flatMap(path => path.points);
      return Math.max(...points.map(point => point.x)) - Math.min(...points.map(point => point.x));
    };
    // Zero is outside the log domain. The visible cone therefore starts with
    // a full ring at axis min rather than a raw-value 90%-scale frustum.
    expect(renderedWidth('cone') / renderedWidth('cylinder')).toBeGreaterThan(0.88);
  });

  it('applies per-point 3-D bar fills and noFill without shaded ghosts', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A', 'B', 'C'],
      threeD: { rotationX: 15, rotationY: 20, shape: 'box' },
      series: [series({
        values: [10, 20, 30],
        dataPointColors: ['FF0000', '00FF00', null],
        dataPointOverrides: [{ idx: 2, fillHidden: true, lineHidden: true }],
      })],
    }), RECT, 1);
    expect(materialFills(rec, 'FF0000').length).toBeGreaterThan(0);
    expect(materialFills(rec, '00FF00').length).toBeGreaterThan(0);
    expect(materialFills(rec, '4472C4')).toHaveLength(0);
  });

  it('preserves a per-point 3-D bar outline dash and point fill over series noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        color: '00000000',
        values: [10],
        dataPointOverrides: [{
          idx: 0,
          color: 'FF0000',
          lineColor: '00FF00',
          lineWidthEmu: 25_400,
          lineDash: 'dash',
        }],
      })],
    }), RECT, 1);
    const outlines = rec.filledPaths.filter(path => path.fillStyle === '#00FF00');
    expect(outlines.length).toBeGreaterThan(0);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#00FF00'
    )).toBe(false);
  });

  it('inherits the authored series dash for 3-D bar faces', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10], lineColor: 'FF0000',
        chartexStyle: { lineDash: 'dash', lineCap: 'rnd', lineJoin: 'bevel' },
      })],
    }), RECT, 1);
    const outlines = rec.filledPaths.filter(path => path.fillStyle === '#FF0000');
    expect(outlines.length).toBeGreaterThan(0);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#FF0000'
    )).toBe(false);
  });

  it('uses a line swatch rather than a filled square for a 3-D line legend', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'], showLegend: true,
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ name: 'Trend', values: [1, 2], lineColor: 'FF0000' })],
    }), RECT, 1);
    expect(rec.paintEvents.some(event => event.kind === 'stroke' && event.strokeStyle === '#FF0000')).toBe(true);
    expect(rec.rects.some(rect => rect.fs === '#FF0000')).toBe(false);
  });

  it('measures and wraps a long 3-D side-legend series name without losing words', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar', categories: ['A'], showLegend: true, legendPos: 'r',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ name: 'Super Duper MPG', values: [1] })],
    }), RECT, 1);
    const legendLines = rec.texts.map(text => text.text)
      .filter(text => /Super|Duper|MPG/.test(text));
    expect(legendLines.length).toBe(2);
    expect(legendLines.join(' ')).toContain('Super');
    expect(legendLines.join(' ')).toContain('Duper');
    expect(legendLines.join(' ')).toContain('MPG');
    expect(legendLines.some(text => text.includes('…'))).toBe(false);
  });

  it('routes 3-D legend sides and manual plot/title layout through the shared frame', () => {
    const top = recordingCtx();
    renderChart(top.ctx, baseModel({
      chartType: 'clusteredBar', title: 'Moved', showLegend: true, legendPos: 't',
      titleManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.55, y: 0.02, w: 0.35, h: 0.08,
      },
      plotAreaManualLayout: {
        layoutTarget: 'inner', xMode: 'edge', yMode: 'edge',
        x: 0.35, y: 0.3, w: 0.45, h: 0.45,
      },
      categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({ name: 'Employees', color: 'FF0000', values: [10] })],
    }), RECT, 1);
    const title = top.texts.find(text => text.text === 'Moved');
    const legend = top.texts.find(text => text.text === 'Employees');
    expect(title?.x).toBeGreaterThan(RECT.w * 0.55);
    expect(legend?.y).toBeLessThan(RECT.h * 0.3);
    const authoredFaces = materialFills(top, 'FF0000');
    expect(authoredFaces.length).toBeGreaterThan(0);
  });

  it('projects explicit 3-D line markers and inherits the series line width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10, 20], showMarker: true, markerSymbol: 'circle', markerSize: 7,
        markerFill: 'FFFFFF', markerLine: 'FF0000', lineWidthEmu: 25400,
      })],
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(2);
    expect(rec.paintEvents.filter(event =>
      event.kind === 'stroke' && event.strokeStyle === '#FF0000'
    )).toHaveLength(2);
  });

  it('projects structured 3-D line-marker fills through the shared paint model', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10, 20], showMarker: true, markerSymbol: 'circle', markerSize: 7,
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'DDEEFF' },
          ],
        },
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(2);
    expect(rec.gradients.every(gradient => gradient.stops.length === 2)).toBe(true);
  });

  it('keeps direct point marker paint authoritative over a 3-D series noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10, 20], showMarker: true, markerSymbol: 'circle',
        markerFill: '00000000',
        dataPointOverrides: [{
          idx: 1,
          markerFillPaint: {
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          },
        }],
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(1);
  });

  it('uses the structured 3-D line marker in its legend key', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'], showLegend: true, legendPos: 'r',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        name: 'Gradient marker', values: [10], showMarker: true,
        markerSymbol: 'circle',
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'DDEEFF' },
          ],
        },
      })],
    }), RECT, 1);
    // One gradient belongs to the plotted marker and one to the legend marker.
    expect(rec.gradients).toHaveLength(2);
  });

  it('skips missing 3-D line points instead of inventing zero markers or labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10, null, 30], showMarker: true, markerSymbol: 'circle',
        dataLabelOverrides: [{ idx: 1, text: 'missing must stay absent' }],
      })],
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(2);
    expect(rec.texts.some(text => text.text === 'missing must stay absent')).toBe(false);
  });

  it('normalizes extreme finite 3-D pie values without dropping the chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [Number.MAX_VALUE, Number.MAX_VALUE, Number.NaN] })],
    }), RECT, 1);
    expect(materialFills(rec, '4472C4').length + materialFills(rec, 'ED7D31').length)
      .toBeGreaterThan(0);
  });

  it('normalizes extreme finite 3-D pie view angles before Canvas geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      firstSliceAngle: Number.MAX_VALUE,
      threeD: { rotationX: 15, rotationY: Number.MAX_VALUE },
      series: [series({ values: [1, 2], lineColor: 'FF0000' })],
    }), RECT, 1);
    const vertices = rec.filledPaths
      .filter(path => path.fillStyle === '#FF0000')
      .flatMap(path => path.points);
    expect(vertices.length).toBeGreaterThan(0);
    expect(vertices.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))).toBe(true);
  });

  it('shares 2-D label composition, font color and unboxed leader-line semantics in 3-D', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['Category'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        name: 'Series', values: [5],
        seriesDataLabels: {
          showCatName: true, showSerName: true, showVal: false, showPercent: false,
          fontColor: 'FF0000', showLeaderLines: true, leaderLineColor: '00FF00',
        },
      })],
    }), RECT, 1);
    const label = rec.texts.find(text => text.text === 'Category Series');
    expect(label?.fillStyle).toBe('#FF0000');
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#00FF00'
    )).toBe(true);
  });

  it('uses authored 3-D pie point fills and does not invent white slice borders', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [1, 1], dataPointColors: ['123456', 'ABCDEF'] })],
    }), RECT, 1);
    expect(materialFills(rec, '123456').length).toBeGreaterThan(0);
    expect(materialFills(rec, 'ABCDEF').length).toBeGreaterThan(0);
    expect(rec.paintEvents.some(event => event.kind === 'stroke' && event.strokeStyle === '#FFFFFF')).toBe(false);
  });

  it('uses authored pie3D hPercent as a thickness multiplier, not a scene aspect ratio', () => {
    const render = (heightPercent: number, authored: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'pie',
        threeD: {
          rotationX: 0,
          rotationY: 0,
          perspective: 0,
          heightPercent,
          heightPercentAuthored: authored,
        },
        series: [series({ values: [1], color: '4472C4' })],
      }), RECT, 1);
      const paths = rec.filledPaths
        .filter(path => isMaterialColor(path.fillStyle, '4472C4'))
        .map(path => path.points.map(point => ({ x: point.x, y: point.y })));
      const points = paths.flat();
      return {
        paths,
        spanY: Math.max(...points.map(point => point.y)) - Math.min(...points.map(point => point.y)),
      };
    };

    const omitted = render(50, false);
    const half = render(50, true);
    const full = render(100, true);
    expect(omitted.spanY).toBeCloseTo(full.spanY, 6);
    expect(half.spanY).toBeLessThan(full.spanY);
  });

  it('uses the radial-family zero-yaw default when 3-D pie rotY is omitted', () => {
    const render = (rotationY?: number) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'pie',
        threeD: { rotationX: 30, perspective: 0, ...(rotationY == null ? {} : { rotationY }) },
        series: [series({ values: [1, 2], dataPointColors: ['123456', 'ABCDEF'] })],
      }), RECT, 1);
      return rec.filledPaths.map(path => ({
        fillStyle: path.fillStyle,
        points: path.points.map(point => [point.x, point.y]),
      }));
    };

    expect(render()).toEqual(render(0));
    expect(render()).not.toEqual(render(20));
  });

  it.each([-15, 15, 89])('keeps a %s-degree 3-D pie inside its chart bounds', rotationX => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      threeD: { rotationX, rotationY: 20 },
      series: [series({ values: [1, 2, 3], lineColor: 'FF0000' })],
    }), RECT, 1);
    const vertices = rec.filledPaths
      .filter(path => path.fillStyle === '#FF0000')
      .flatMap(path => path.points);
    expect(vertices.length).toBeGreaterThan(0);
    expect(vertices.every(point =>
      point.x >= RECT.x - 1e-6 && point.x <= RECT.x + RECT.w + 1e-6
      && point.y >= RECT.y - 1e-6 && point.y <= RECT.y + RECT.h + 1e-6
    )).toBe(true);
  });

  it.each([-15, 15, 89])('fits a maximally exploded %s-degree 3-D pie inside its chart bounds', rotationX => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      threeD: { rotationX, rotationY: 20 },
      series: [series({
        values: [1, 2, 3],
        lineColor: 'FF0000',
        dataPointOverrides: [{ idx: 0, explosion: 100 }],
      })],
    }), RECT, 1);
    const vertices = rec.filledPaths
      .filter(path => path.fillStyle === '#FF0000')
      .flatMap(path => path.points);
    expect(vertices.length).toBeGreaterThan(0);
    expect(vertices.every(point =>
      point.x >= RECT.x - 1e-6 && point.x <= RECT.x + RECT.w + 1e-6
      && point.y >= RECT.y - 1e-6 && point.y <= RECT.y + RECT.h + 1e-6
    )).toBe(true);
  });

  it('clamps the 3-D bar datum to a positive logarithmic value-axis domain', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      valMin: 1,
      valMax: 1000,
      valAxisLogBase: 10,
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [10], lineColor: 'FF0000' })],
    }), RECT, 1);
    const vertices = rec.filledPaths
      .filter(path => path.fillStyle === '#FF0000')
      .flatMap(path => path.points);
    expect(vertices.length).toBeGreaterThan(0);
    expect(vertices.every(point =>
      Number.isFinite(point.x) && Number.isFinite(point.y)
      && point.x >= RECT.x - 1e-6 && point.x <= RECT.x + RECT.w + 1e-6
      && point.y >= RECT.y - 1e-6 && point.y <= RECT.y + RECT.h + 1e-6
    )).toBe(true);
  });

  it('honors per-point 3-D pie labels and percent formatting without the chart-wide flag', () => {
    const rec = recordingCtx(() => 18);
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      showDataLabels: false,
      categories: ['First', 'Second'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [1, 3],
        categories: ['First', 'Second'],
        dataLabelOverrides: [{
          idx: 0,
          text: '',
          showVal: false,
          showCatName: true,
          showPercent: true,
          formatCode: '0.0%',
          separator: ' / ',
        }],
      })],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('First / 25.0%');
    expect(rec.texts.map(item => item.text)).not.toContain('Second');
  });

  it('draws pie leader lines only for labels placed outside the slice', () => {
    const inside = recordingCtx(() => 8);
    renderChart(inside.ctx, baseModel({
      chartType: 'pie', threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [1, 1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          showLeaderLines: true, leaderLineColor: '00FF00', position: 'ctr',
        },
      })],
    }), RECT, 1);
    expect(inside.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#00FF00'
    )).toBe(false);

    const outside = recordingCtx(() => 200);
    renderChart(outside.ctx, baseModel({
      chartType: 'pie', threeD: { rotationX: 15, rotationY: 20 },
      chartStyleRoles: { leaderLine: { lineColors: ['FF0000'] } },
      series: [series({
        values: [99, 1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          showLeaderLines: true, leaderLineColor: '00FF00',
        },
      })],
    }), RECT, 1);
    expect(outside.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#00FF00'
    )).toBe(true);
    expect(outside.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#FF0000'
    )).toBe(false);
  });

  it('uses linked leaderLine paint when the data-label block omits line paint', () => {
    const render = (lineHidden: boolean) => {
      const rec = recordingCtx(() => 200);
      renderChart(rec.ctx, baseModel({
        chartType: 'pie', threeD: { rotationX: 15, rotationY: 20 },
        chartStyleRoles: {
          leaderLine: { lineColors: ['CC5500'], lineWidthEmu: 19050, lineHidden },
        },
        series: [series({
          values: [99, 1],
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
            showLeaderLines: true,
          },
        })],
      }), RECT, 1);
      return rec;
    };
    const visible = render(false);
    expect(visible.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#CC5500'
    )).toBe(true);
    expect(render(true).paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#CC5500'
    )).toBe(false);
  });

  it('keeps non-finite firstSliceAngle out of 3-D pie geometry', () => {
    for (const firstSliceAngle of [Number.NaN, Number.POSITIVE_INFINITY]) {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'pie', categories: ['A', 'B'], firstSliceAngle,
        threeD: { rotationX: 15, rotationY: 20 },
        series: [series({ values: [1, 2], lineColor: 'FF0000' })],
      }), RECT, 1);
      const outline = rec.filledPaths.filter(path => path.fillStyle === '#FF0000');
      expect(outline.length).toBeGreaterThan(0);
      expect(outline.flatMap(path => path.points)
        .every(point => Number.isFinite(point.x) && Number.isFinite(point.y))).toBe(true);
    }
  });

  it('applies per-point 3-D pie outline and explosion authoring', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [1, 1],
        dataPointOverrides: [{
          idx: 0, color: 'FF0000', lineColor: '00FF00', lineWidthEmu: 25_400,
          lineDash: 'dash', explosion: 20,
        }],
      })],
    }), RECT, 1);
    const outlines = rec.filledPaths.filter(path => path.fillStyle === '#00FF00');
    expect(outlines.length).toBeGreaterThan(0);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#00FF00'
    )).toBe(false);
  });

  it('paints each identical shared 3-D pie outline primitive only once', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: ['A', 'B', 'C', 'D'],
      threeD: { rotationX: 30, rotationY: 0, perspective: 0 },
      series: [series({
        values: [8, 6, 2, 2],
        lineColor: '000000',
        lineWidthEmu: 12_700,
      })],
    }), RECT, 1);
    const outlines = rec.filledPaths.filter(path => path.fillStyle === '#000000');
    const geometryKey = (path: (typeof outlines)[number]) => path.points
      .map(point => `${point.x.toFixed(6)},${point.y.toFixed(6)}`)
      .sort()
      .join('|');
    expect(outlines.length).toBeGreaterThan(0);
    expect(new Set(outlines.map(geometryKey)).size).toBe(outlines.length);
    const lastMaterialFill = rec.paintEvents.reduce((last, event, index) =>
      event.kind === 'fill' && ['4472C4', 'ED7D31', '70AD47', 'A5A5A5'].some(color =>
        isMaterialColor(event.fillStyle, color)
      ) ? index : last, -1);
    const firstOutlineFill = rec.paintEvents.findIndex(event =>
      event.kind === 'fill' && event.fillStyle === '#000000');
    const verticalOutlinePrimitives = outlines.filter(path => {
      const xs = path.points.map(point => point.x);
      const ys = path.points.map(point => point.y);
      return Math.max(...ys) - Math.min(...ys) > 8
        && Math.max(...xs) - Math.min(...xs) < 4;
    });
    expect(lastMaterialFill).toBeGreaterThanOrEqual(0);
    expect(firstOutlineFill).toBeGreaterThan(lastMaterialFill);
    expect(verticalOutlinePrimitives.length).toBeGreaterThanOrEqual(3);
  });

  it('paints authored 3-D axis titles after tick labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      catAxisTitle: 'Categories', valAxisTitle: 'Employees',
      catAxisTitleFontSizeHpt: 1000, valAxisTitleFontSizeHpt: 1000,
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values: [1, 2] })],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('Categories');
    expect(rec.texts.map(item => item.text)).toContain('Employees');
  });

  it('reserves the effective DrawingML text insets for a 3-D axis title', () => {
    const titleY = (withInsets: boolean): number => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line', categories: ['A', 'B'],
        catAxisTitle: 'Categories', catAxisTitleFontSizeHpt: 1000,
        catAxisTitleTextVerticalInsetEmu: withInsets ? 91_440 : undefined,
        threeD: { rotationX: 15, rotationY: 20 },
        series: [series({ values: [1, 2] })],
      }), RECT, 1);
      expect(rec.texts.map(item => item.text)).toContain('Categories');
      return rec.translations.at(-1)?.y ?? Number.NaN;
    };

    expect(titleY(true)).toBeCloseTo(titleY(false) - 7.2);
  });

  it('rejects derived 3-D face work before expanding a moderate source cache', () => {
    const rec = recordingCtx();
    const values = Array.from({ length: 300 }, (_, index) => index + 1);
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: values.map(String),
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({ values })],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).toContain('(too many data points)');
    expect(rec.paintEvents.filter(event => event.kind === 'fill')).toHaveLength(0);
  });

  it('does not charge ordinary 3-D area data for worst-case clipping expansion', () => {
    const count = 334;
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: Array.from({ length: count }, (_, index) => 40 + 10 * Math.sin(index / 12)),
        lineHidden: true,
      })],
    }), RECT, 1);
    expect(rec.texts.map(item => item.text)).not.toContain('(too many data points)');
    expect(materialFills(rec, '4472C4').length).toBeGreaterThan(0);
  });

  it('bounds projected dash primitives cumulatively across 3-D series', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20 },
      series: [
        series({ values: [0, 1], chartexStyle: { lineDash: 'dash' } }),
        series({ values: [1, 0], chartexStyle: { lineDash: 'dash' } }),
        series({ values: [0.2, 0.8], chartexStyle: { lineDash: 'dash' } }),
        series({ values: [0.8, 0.2], chartexStyle: { lineDash: 'dash' } }),
      ],
    }), { x: 0, y: 0, w: 200_000, h: 200_000 }, 1);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
    expect(rec.rects.some(rect => rect.fs === '#FFFFFF')).toBe(false);
  });

  it('bounds expanded authored mesh outlines across many 3-D data points', () => {
    const count = 277; // 277 * static weight 36 = 9,972: passes source preflight.
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      threeD: { rotationX: 15, rotationY: 20, shape: 'cylinder' },
      series: [series({
        values: new Array(count).fill(1),
        lineColor: 'FF0000',
        chartexStyle: { lineDash: 'dash' },
      })],
    }), { x: 0, y: 0, w: 100_000, h: 100_000 }, 1);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it.each(['clusteredBar', 'line'] as const)(
    'keeps authored custom data labels on projected 3-D %s points',
    chartType => {
      const rec = recordingCtx(() => 30);
      renderChart(rec.ctx, baseModel({
        chartType, categories: ['A'], showDataLabels: false,
        threeD: { rotationX: 15, rotationY: 20 },
        series: [series({
          values: [10],
          dataLabelOverrides: [{ idx: 0, text: 'Projected label', position: 't' }],
        })],
      }), RECT, 1);
      const label = rec.texts.find(text => text.text === 'Projected label');
      expect(label).toBeDefined();
      expect(Number.isFinite(label?.x)).toBe(true);
      expect(Number.isFinite(label?.y)).toBe(true);
    },
  );

  it('paints bounded rich runs and callout authoring on a projected 3-D label', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'], showDataLabels: false,
      threeD: { rotationX: 15, rotationY: 20 },
      series: [series({
        values: [10],
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: false,
          showLeaderLines: true, leaderLineColor: '808080',
        },
        dataLabelOverrides: [{
          idx: 0, text: 'Rich label', position: 'outEnd',
          richRuns: [
            { text: 'Rich ', color: 'FF0000', fontSizeHpt: 1200, bold: true },
            { text: 'label', color: '0000FF', fontSizeHpt: 1000, bold: false },
          ],
          labelBox: { fill: 'FFFFFF', borderColor: '000000', borderWidthEmu: 12_700 },
        }],
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === 'Rich ' && text.fillStyle === '#FF0000')).toBe(true);
    expect(rec.texts.some(text => text.text === 'label' && text.fillStyle === '#0000FF')).toBe(true);
    expect(rec.rects.some(rect => rect.fs === '#FFFFFF')).toBe(true);
    expect(rec.strokeRects.some(rect => rect.ss === '#000000')).toBe(true);
  });
});

describe('chart drawing user-shape text boxes', () => {
  it('applies DrawingML default text insets inside chart user shapes', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0,
        y: 0,
        w: 1,
        h: 0.2,
        paragraphs: [{ runs: [{ text: 'Title', fontSizeHpt: 1000 }] }],
      }],
    }), RECT, 1);

    const title = rec.texts.find(text => text.text === 'Title');
    expect(title?.x).toBeCloseTo(7.2, 6);
    expect(title?.y).toBeCloseTo(3.6 + 9, 6);
  });

  it('honors explicit asymmetric DrawingML text insets and the content-box alignment', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0.1,
        y: 0.1,
        w: 0.5,
        h: 0.2,
        lIns: 12700,
        tIns: 25400,
        rIns: 38100,
        bIns: 50800,
        paragraphs: [{ align: 'r', runs: [{ text: 'Right', fontSizeHpt: 1000 }] }],
      }],
    }), RECT, 1);

    const text = rec.texts.find(item => item.text === 'Right');
    const contentRight = RECT.w * 0.6 - 3;
    expect((text?.x ?? 0) + (text?.width ?? 0)).toBeCloseTo(contentRight, 6);
    expect(text?.y).toBeCloseTo(RECT.h * 0.1 + 2 + 9, 6);
  });

  it('draws relative paragraphs with authored run formatting above the chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0.1,
        y: 0.05,
        w: 0.8,
        h: 0.15,
        verticalAnchor: 'b',
        paragraphs: [{
          align: 'ctr',
          runs: [
            { text: 'Authored ', fontSizeHpt: 2000, bold: true, color: '1696D2', fontFace: 'Lato' },
            { text: 'title', fontSizeHpt: 1200, fontFace: 'Arial' },
          ],
        }],
      }],
    }), RECT, 1);

    const authored = rec.texts.find(text => text.text === 'Authored ');
    const suffix = rec.texts.find(text => text.text === 'title');
    expect(authored).toBeDefined();
    expect(suffix).toBeDefined();
    expect(authored?.font).toContain('bold 20px');
    expect(authored?.font).toContain('Lato');
    expect(authored?.fillStyle).toBe('#1696D2');
    expect(suffix?.font).toContain('12px');
    expect(suffix?.font).toContain('Arial');
    expect(authored?.x).toBeGreaterThan(RECT.w * 0.1);
    expect((suffix?.x ?? 0)).toBeGreaterThan(authored?.x ?? 0);
    expect(authored?.y).toBeGreaterThan(RECT.h * 0.05);
    expect(authored?.y).toBeLessThanOrEqual(RECT.h * 0.2);
  });

  it('wraps DrawingML text inside its authored rectangle unless wrap is none', () => {
    const wrapped = recordingCtx();
    renderChart(wrapped.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0,
        y: 0,
        w: 0.16,
        h: 0.3,
        paragraphs: [{ runs: [{ text: 'Alpha beta gamma', fontSizeHpt: 1200 }] }],
      }],
    }), RECT, 1);

    const wrappedWords = wrapped.texts.filter(text => ['Alpha', 'beta', 'gamma'].includes(text.text));
    expect(wrappedWords).toHaveLength(3);
    expect(new Set(wrappedWords.map(text => text.y)).size).toBeGreaterThan(1);

    const unwrapped = recordingCtx();
    renderChart(unwrapped.ctx, baseModel({
      categories: ['A'],
      series: [series({ values: [1] })],
      chartTextBoxes: [{
        x: 0,
        y: 0,
        w: 0.16,
        h: 0.3,
        wrap: 'none',
        paragraphs: [{ runs: [{ text: 'Alpha beta gamma', fontSizeHpt: 1200 }] }],
      }],
    }), RECT, 1);
    expect(unwrapped.texts.some(text => text.text === 'Alpha beta gamma')).toBe(true);
  });
});

describe('bar chart authored layout and fills', () => {
  it('honors a manually positioned title and ignores its authored width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      title: 'Readiness score',
      titleManualLayout: { x: 0.13, y: 0.03, w: 0.5, h: 0.08, xMode: 'edge', yMode: 'edge' },
      categories: ['A'],
      series: [series({ name: 'S', values: [1] })],
    }), RECT, 1);
    const title = rec.texts.find(text => text.text === 'Readiness score');
    expect(title).toBeDefined();
    expect(title?.x).toBeCloseTo(RECT.w * 0.13 + (title?.width ?? 0) / 2, 4);
    expect(title?.x).toBeLessThan(RECT.w / 2);
  });

  it('keeps automatic title width when manual layout supplies only an edge position', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      title: 'Readiness score by pillar',
      titleManualLayout: { x: 0.13, y: 0.03, xMode: 'edge', yMode: 'edge' },
      categories: ['A'],
      series: [series({ name: 'S', values: [1] })],
    }), RECT, 1);

    const title = rec.texts.find(text => text.text === 'Readiness score by pillar');
    expect(title).toBeDefined();
    expect(title?.x).toBeCloseTo(RECT.w * 0.13 + (title?.width ?? 0) / 2, 4);
  });

  it('uses a series pattern fill for bars and legend swatches', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        name: 'Patterned',
        color: '111111',
        values: [1],
        fillPattern: { fillType: 'pattern', fg: '777777', bg: 'FFFFFF', preset: 'pct30' },
        lineColor: '595959',
        lineWidthEmu: 12700,
      })],
      showLegend: true,
      legendPos: 'r',
    }), RECT, 1);
    // A headless test has no bitmap canvas, so resolveFill deliberately falls
    // back to the pattern foreground. Both the bar and key must use it.
    expect(rec.rects.filter(rect => rect.fs === 'rgba(119,119,119,1)').length).toBeGreaterThanOrEqual(2);
    // The authored series outline applies to both the plotted bars and their
    // matching legend key. Excel's patterned key is not borderless.
    expect(rec.strokeRects.filter(rect => rect.ss === '#595959' && rect.lw === 1)).toHaveLength(2);
  });

  it('keeps a filled legend key at the Office-observed 7pt square and preserves its outline', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        name: 'Outlined',
        color: 'FF00FF',
        values: [1],
        lineColor: '000000',
        lineWidthEmu: 12700,
      })],
      showLegend: true,
      legendPos: 'r',
      legendFontSizeHpt: 1500,
    }), RECT, 1);

    const key = rec.rects.find(rect =>
      rect.fs === '#FF00FF' && Math.abs(rect.w - rect.h) < 0.01 && rect.w < 10
    );
    expect(key).toMatchObject({ w: 7, h: 7 });
    expect(rec.strokeRects.some(rect =>
      rect.ss === '#000000'
      && rect.lw === 1
      && Math.abs(rect.w - 6) < 0.01
      && Math.abs(rect.h - 6) < 0.01
    )).toBe(true);
  });

  it('keeps a short top-legend label intact at fractional display metrics', () => {
    // Aptos Narrow at 13.3333px produces this fractional width in Chromium.
    // Adding a 7pt key and then subtracting it again loses one ULP; without a
    // bounded measurement epsilon that false overflow turns `disp` into `di…`.
    const rec = recordingCtx((text, fontPx) => {
      if (Math.abs(fontPx - 13.3333) > 0.001) return null;
      if (text === 'disp') return 24.453475952148438;
      if (text === 'di…') return 23.704971313476562;
      if (text === 'dis…') return 30.369964599609375;
      return null;
    });
    renderChart(rec.ctx, baseModel({
      chartType: 'funnel',
      title: 'Sales Funnel',
      categories: ['1', '2', '3', '4', '5'],
      series: [series({ name: 'disp', values: [3, 3, 2, 4, 5] })],
      showLegend: true,
      legendPos: 't',
    }), { x: 0, y: 0, w: 457.2, h: 292.608 }, 4 / 3);

    expect(rec.texts.some(text => text.text === 'disp')).toBe(true);
    expect(rec.texts.some(text => text.text.includes('…'))).toBe(false);
  });

  it('wraps a measured top legend into centered in-bounds rows without changing authored text style', () => {
    const names = Array.from({ length: 12 }, (_, index) =>
      `Series ${String(index + 1).padStart(2, '0')} alpha`
    );
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: names.map((name, index) => series({ name, values: [index + 1] })),
      showLegend: true,
      legendPos: 't',
      legendFontSizeHpt: 1200,
      legendFontBold: true,
      legendFontFace: 'Legend Face',
    }), RECT, 1);

    const labels = rec.texts.filter(text => text.text.startsWith('Series '));
    expect(labels.map(label => label.text)).toEqual(names);
    expect(new Set(labels.map(label => Math.round(label.y))).size).toBeGreaterThan(1);
    expect(labels.every(label =>
      label.x >= RECT.x
      && label.x + (label.width ?? 0) <= RECT.x + RECT.w
      && label.y >= RECT.y
      && label.y <= RECT.y + RECT.h
    )).toBe(true);
    expect(labels.every(label =>
      label.font?.includes('bold 12px') && label.font.includes('Legend Face')
    )).toBe(true);
  });

  it('uses the painted top-legend content width when reserving wrapped rows', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [
        series({ name: 'AAAAA', values: [1] }),
        series({ name: 'BBBBB', values: [2] }),
      ],
      showLegend: true,
      legendPos: 't',
      legendFontSizeHpt: 1000,
    }), { x: 0, y: 0, w: 96, h: 200 }, 1);

    // The fixed 7pt keys leave the pair wider than the actual w - 8 content
    // rectangle, so both rows must be reserved and painted from the same plan.
    const labels = rec.texts.filter(text => text.text === 'AAAAA' || text.text === 'BBBBB');
    expect(labels.map(label => label.text)).toEqual(['AAAAA', 'BBBBB']);
    expect(labels[1].y).toBeGreaterThan(labels[0].y);
  });

  it('keeps a long category-driven side legend to complete, non-overlapping rows inside the chart', () => {
    const categories = Array.from({ length: 20 }, (_, index) =>
      `Category ${String(index + 1).padStart(2, '0')} with a deliberately long label`
    );
    const rect = { x: 0, y: 0, w: 320, h: 180 };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories,
      series: [series({ values: categories.map((_, index) => index + 1) })],
      showLegend: true,
      legendPos: 'r',
      legendFontSizeHpt: 1000,
    }), rect, 1);

    const labels = rec.texts.filter(text => text.text.startsWith('Category '));
    expect(labels.length).toBeGreaterThan(0);
    expect(labels.length).toBeLessThan(categories.length);
    const ys = labels.map(label => label.y);
    expect(ys.every((value, index) => index === 0 || value - ys[index - 1] >= 13.9)).toBe(true);
    expect(labels.every(label =>
      label.x >= rect.x
      && label.x + (label.width ?? 0) <= rect.x + rect.w
      && label.y >= rect.y
      && label.y <= rect.y + rect.h
    )).toBe(true);
    expect(labels.every(label => label.text.endsWith('…'))).toBe(true);
  });

  it('uses chart categories when an empty pie-series category cache would under-measure top rows', () => {
    const categories = ['Category Alpha', 'Category Beta'];
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories,
      series: [series({ values: [1, 2], categories: [] })],
      showLegend: true,
      legendPos: 't',
      legendFontSizeHpt: 1000,
    }), { x: 0, y: 0, w: 200, h: 200 }, 1);

    const labels = rec.texts.filter(text => categories.includes(text.text));
    expect(labels.map(label => label.text)).toEqual(categories);
    expect(labels[1].y).toBeGreaterThan(labels[0].y);
  });

  it('measures side-legend width from fallback chart categories before elision', () => {
    const categories = ['Category Alpha Long', 'Category Beta Long'];
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'doughnut',
      categories,
      series: [series({ values: [1, 2], categories: [] })],
      showLegend: true,
      legendPos: 'r',
      legendFontSizeHpt: 1000,
    }), RECT, 1);

    const labels = rec.texts
      .map(text => text.text)
      .filter(text => text.startsWith('Category '));
    expect(labels).toEqual(categories);
  });

  it('does not paint an automatic side-legend key outside a very narrow chart', () => {
    const rect = { x: 0, y: 0, w: 20, h: 200 };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['A'],
      series: [series({ values: [1] })],
      showLegend: true,
      legendPos: 'r',
    }), rect, 1);

    expect(rec.rects.every(item =>
      item.x >= rect.x
      && item.x + item.w <= rect.x + rect.w
      && item.y >= rect.y
      && item.y + item.h <= rect.y + rect.h
    )).toBe(true);
  });

  it('keeps a valid manual legend rectangle authoritative over automatic side packing', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: ['Manual A', 'Manual B', 'Manual C'].map((name, index) =>
        series({ name, values: [index + 1] })
      ),
      showLegend: true,
      legendPos: 'r',
      legendManualLayout: {
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.1, y: 0.1, w: 0.5, h: 0.15,
      },
    }), RECT, 1);

    const labels = rec.texts.filter(text => text.text.startsWith('Manual '));
    expect(labels).toHaveLength(3);
    expect(labels.every(label =>
      label.x >= RECT.w * 0.1
      && label.x + (label.width ?? 0) <= RECT.w * 0.6
      && label.y >= RECT.h * 0.1
      && label.y <= RECT.h * 0.25
    )).toBe(true);
  });

  it('keeps a non-overlay manual top legend outside the automatic bar plot', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      title: 'Revenue',
      titlePresent: true,
      titleFontSizeHpt: 1200,
      categories: ['A'],
      series: [series({ name: 'Series', values: [10] })],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 5,
      valAxisFontSizeHpt: 1000,
      showLegend: true,
      legendPos: 't',
      legendOverlay: false,
      legendFillColor: '123456',
      legendFontSizeHpt: 1000,
      legendManualLayout: {
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.1, y: 0.25, w: 0.5, h: 0.1,
      },
    }), RECT, 1);

    const legend = rec.rects.find(rect => rect.fs === '#123456');
    if (!legend) throw new Error('expected the authored legend frame');
    const horizontalGridLines = rec.strokedPaths
      .filter(path => path.points.length === 2
        && Math.abs(path.points[0].y - path.points[1].y) < 0.001
        && Math.abs(path.points[1].x - path.points[0].x) > RECT.w / 2);
    expect(horizontalGridLines.length).toBeGreaterThan(0);
    const plotTop = Math.min(...horizontalGridLines.map(path => path.points[0].y));
    expect(plotTop).toBeGreaterThanOrEqual(legend.y + legend.h);
  });

  it('does not reserve plot space for an overlay legend and retains manual placement', () => {
    const render = (overlay: boolean) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar',
        categories: ['A'],
        series: [series({ name: 'Overlay', color: 'FF0000', values: [10] })],
        showLegend: true,
        legendPos: 'r',
        legendOverlay: overlay,
        legendManualLayout: {
          xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
          x: 0.55, y: 0.1, w: 0.35, h: 0.2,
        },
      }), RECT, 1);
      return rec;
    };
    const reserved = render(false);
    const overlay = render(true);
    const widestRed = (rec: Recorded) => Math.max(
      ...rec.rects.filter(rect => rect.fs === '#FF0000').map(rect => rect.w),
    );
    expect(widestRed(overlay)).toBeGreaterThan(widestRed(reserved));
    const legend = overlay.texts.find(text => text.text === 'Overlay');
    expect(legend?.x).toBeGreaterThanOrEqual(RECT.w * 0.55);
    expect(legend?.x).toBeLessThanOrEqual(RECT.w * 0.9);
  });

  it('applies indexed legend deletion and entry-local text properties without reordering', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: ['First', 'Deleted', 'Styled'].map((name, index) =>
        series({ name, values: [index + 1] })
      ),
      showLegend: true,
      legendPos: 'r',
      legendEntries: [
        { idx: 1, deleted: true },
        {
          idx: 2,
          fontFace: 'Entry Face',
          fontColor: 'AABBCC',
          fontSizeHpt: 1400,
          fontBold: true,
        },
      ],
    }), RECT, 1);

    const labels = rec.texts.filter(text => ['First', 'Deleted', 'Styled'].includes(text.text));
    expect(labels.map(label => label.text)).toEqual(['First', 'Styled']);
    expect(labels[1]).toMatchObject({ fillStyle: '#AABBCC' });
    expect(labels[1].font).toContain('bold 14px');
    expect(labels[1].font).toContain('Entry Face');
  });

  it('indexes legend-entry overrides against point-driven pie entries', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Alpha', 'Beta', 'Gamma'],
      series: [series({ values: [2, 3, 5] })],
      showLegend: true,
      legendPos: 'r',
      legendEntries: [
        { idx: 0, deleted: false },
        { idx: 1, deleted: true },
        { idx: 2, fontColor: '008800', fontBold: true },
      ],
    }), RECT, 1);

    const labels = rec.texts.filter(text => ['Alpha', 'Beta', 'Gamma'].includes(text.text));
    expect(labels.map(label => label.text)).toEqual(['Alpha', 'Gamma']);
    expect(labels[1]).toMatchObject({ fillStyle: '#008800' });
    expect(labels[1].font).toContain('bold');
  });

  it('applies the same indexed legend overrides to a 3-D chart', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: ['First 3D', 'Deleted 3D', 'Styled 3D'].map((name, index) =>
        series({ name, values: [index + 1] })
      ),
      showLegend: true,
      legendPos: 'r',
      legendEntries: [
        { idx: 1, deleted: true },
        { idx: 2, fontColor: 'CC00CC', fontSizeHpt: 1300, fontBold: true },
      ],
      threeD: { rotationX: 15, rotationY: 20 },
    }), RECT, 1);

    const labels = rec.texts.filter(text => text.text.endsWith('3D'));
    expect(labels.map(label => label.text)).toEqual(['First 3D', 'Styled 3D']);
    expect(labels[1]).toMatchObject({ fillStyle: '#CC00CC' });
    expect(labels[1].font).toContain('bold 13px');
  });

  it('paints an authored legend-frame fill and outline behind its manual content box', () => {
    const rec = recordingCtx();
    const chart = baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({ name: 'Framed', values: [1] })],
      showLegend: true,
      legendPos: 'r',
      legendManualLayout: {
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.1, y: 0.2, w: 0.5, h: 0.3,
      },
    });
    Object.assign(chart, {
      legendFillColor: 'FFFFFF',
      legendLineColor: '808080',
      legendLineWidthEmu: 3175,
      legendLineDash: 'dot',
      legendLineCap: 'rnd',
      legendLineJoin: 'round',
    });

    renderChart(rec.ctx, chart, RECT, 1);

    expect(rec.rects).toContainEqual({ x: 64, y: 72, w: 320, h: 108, fs: '#FFFFFF' });
    expect(rec.strokeRects).toContainEqual({
      x: 64.25, y: 72.25, w: 319.5, h: 107.5,
      ss: '#808080', lw: 0.5, dash: [0.75, 1.5], cap: 'round', join: 'round',
    });
  });

  it('uses linked legend paint only when direct frame paint is omitted', () => {
    const linked = recordingCtx();
    const chart = baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({ name: 'Linked', values: [1] })],
      showLegend: true,
      legendPos: 'r',
      legendManualLayout: {
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.1, y: 0.2, w: 0.5, h: 0.3,
      },
      chartStyleRoles: {
        legend: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          lineColors: ['808080'],
          lineWidthEmu: 3175,
          lineCustomDash: [{ dash: 1.25, space: 0.75 }],
          lineCompound: 'dbl',
          lineCap: 'sq',
          lineJoin: 'bevel',
        },
      },
    });
    renderChart(linked.ctx, chart, RECT, 1, 30);
    expect(linked.gradients).toHaveLength(1);
    expect(linked.rects).toContainEqual({ x: 64, y: 72, w: 320, h: 108, fs: '[object Object]' });
    expect(linked.strokeRects.filter(rect => rect.ss === '#808080')).toEqual([
      {
        x: 64 + 1 / 12, y: 72 + 1 / 12, w: 320 - 1 / 6, h: 108 - 1 / 6,
        ss: '#808080', lw: 1 / 6, dash: [0.625, 0.375], cap: 'square', join: 'bevel',
      },
      {
        x: 64 + 5 / 12, y: 72 + 5 / 12, w: 320 - 5 / 6, h: 108 - 5 / 6,
        ss: '#808080', lw: 1 / 6, dash: [0.625, 0.375], cap: 'square', join: 'bevel',
      },
    ]);

    const directLineGeometry = recordingCtx();
    renderChart(directLineGeometry.ctx, {
      ...chart,
      legendLineDash: 'solid',
      legendLineCap: 'rnd',
      legendLineJoin: 'round',
    }, RECT, 1);
    expect(directLineGeometry.strokeRects).toContainEqual(expect.objectContaining({
      ss: '#808080', dash: [], cap: 'round', join: 'round',
    }));

    const directEmptyDash = recordingCtx();
    renderChart(directEmptyDash.ctx, { ...chart, legendLineCustomDash: [] }, RECT, 1);
    expect(directEmptyDash.strokeRects.find(rect => rect.ss === '#808080')?.dash)
      .toEqual([]);

    const directNoFill = recordingCtx();
    renderChart(directNoFill.ctx, {
      ...chart,
      legendFillHidden: true,
      legendFillPaintAuthored: true,
      legendLineHidden: true,
      legendLinePaintAuthored: true,
    }, RECT, 1);
    expect(directNoFill.rects).not.toContainEqual(
      expect.objectContaining({ x: 64, y: 72, w: 320, h: 108 }),
    );
    expect(directNoFill.strokeRects).not.toContainEqual(
      expect.objectContaining({ x: 64.25, y: 72.25, w: 319.5, h: 107.5 }),
    );
  });

  it('uses linked plot-area structured fill only when direct paint is omitted', () => {
    const chart = baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      chartStyleRoles: {
        plotArea: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 45,
            rotWithShape: false,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          lineColors: ['445566'],
          lineWidthEmu: 9525,
          lineCustomDash: [{ dash: 1.25, space: 0.75 }],
          lineCompound: 'dbl',
          lineCap: 'sq',
          lineJoin: 'bevel',
        },
      },
    });

    const linked = recordingCtx();
    renderChart(linked.ctx, chart, RECT, 1, 30);
    expect(linked.gradients).toHaveLength(1);
    expect(linked.gradients[0]?.stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(221,238,255,1)' },
    ]);
    expect(linked.rects).toContainEqual(expect.objectContaining({ fs: '[object Object]' }));
    expect(linked.strokeRects.filter(rect => rect.ss === '#445566')).toEqual([
      expect.objectContaining({ lw: 0.25, cap: 'square', join: 'bevel' }),
      expect.objectContaining({ lw: 0.25, cap: 'square', join: 'bevel' }),
    ]);
    expect(linked.strokeRects.find(rect => rect.ss === '#445566')?.dash)
      .toEqual([0.9375, 0.5625]);

    const automaticHostFallback = recordingCtx();
    renderChart(automaticHostFallback.ctx, {
      ...chart,
      plotAreaBg: 'FFFFFF',
      plotAreaFillAutomatic: true,
    }, RECT, 1, 30);
    expect(automaticHostFallback.gradients).toHaveLength(1);
    expect(automaticHostFallback.gradients[0]?.stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(221,238,255,1)' },
    ]);

    const compatibilityDirect = recordingCtx();
    renderChart(compatibilityDirect.ctx, { ...chart, plotAreaBg: 'FFFFFF' }, RECT, 1, 30);
    expect(compatibilityDirect.gradients).toHaveLength(0);

    const directEmptyDash = recordingCtx();
    renderChart(directEmptyDash.ctx, { ...chart, plotAreaLineCustomDash: [] }, RECT, 1, 30);
    expect(directEmptyDash.strokeRects.find(rect => rect.ss === '#445566')?.dash)
      .toEqual([]);

    const directNoFill = recordingCtx();
    renderChart(directNoFill.ctx, {
      ...chart,
      plotAreaFillHidden: true,
      plotAreaFillPaintAuthored: true,
      plotAreaLineHidden: true,
      plotAreaLinePaintAuthored: true,
    }, RECT, 1, 30);
    expect(directNoFill.gradients).toHaveLength(0);
    expect(directNoFill.rects).not.toContainEqual(expect.objectContaining({ fs: '[object Object]' }));
    expect(directNoFill.strokeRects).not.toContainEqual(expect.objectContaining({ ss: '#445566' }));
  });

  it('strokes chart, plot, and legend frames with structured DrawingML line paint', () => {
    const linePaint = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 30,
      stops: [
        { position: 0, color: '112233' },
        { position: 1, color: 'DDEEFF' },
      ],
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      showLegend: true,
      legendPos: 'r',
      chartStyleRoles: {
        chartArea: { linePaints: [linePaint], lineWidthEmu: 9525 },
        plotArea: { linePaints: [linePaint], lineWidthEmu: 9525 },
        legend: { linePaints: [linePaint], lineWidthEmu: 9525 },
      },
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(3);
    expect(rec.strokeRects.filter(rect => rect.ss === '[object Object]')).toHaveLength(3);
  });

  it('keeps a bare direct preset-dash choice authoritative over linked dash', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'], series: [series({ values: [1, 2] })],
      plotAreaLineColor: '445566',
      plotAreaLineDashAuthored: true,
      chartStyleRoles: { plotArea: { lineDash: 'dash' } },
    }), RECT, 1);
    expect(rec.strokeRects.find(rect => rect.ss === '#445566')?.dash).toEqual([]);
  });

  it('honors legend manual-layout x/y when width and height are omitted', () => {
    const common = baseModel({
      chartType: 'line', categories: ['A'], series: [series({ values: [1] })],
      showLegend: true, legendPos: 'r', legendFillColor: '123456',
    });
    const automatic = recordingCtx();
    renderChart(automatic.ctx, common, RECT, 1);
    const manual = recordingCtx();
    renderChart(manual.ctx, {
      ...common,
      legendManualLayout: {
        xMode: 'factor', yMode: 'factor',
        x: 0.1, y: 0.2,
      },
    }, RECT, 1);
    const autoFrame = automatic.rects.find(rect => rect.fs === '#123456');
    const manualFrame = manual.rects.find(rect => rect.fs === '#123456');
    expect(autoFrame).toBeDefined();
    expect(manualFrame).toEqual({
      ...autoFrame as NonNullable<typeof autoFrame>,
      x: (autoFrame as NonNullable<typeof autoFrame>).x + RECT.w * 0.1,
      y: (autoFrame as NonNullable<typeof autoFrame>).y + RECT.h * 0.2,
    });
  });

  it.each([
    ['pie', baseModel({
      chartType: 'pie', categories: ['A', 'B'], series: [series({ values: [1, 2] })],
    })],
    ['radar', baseModel({
      chartType: 'radar', categories: ['A', 'B', 'C'],
      series: [series({ values: [1, 2, 3] })],
    })],
    ['surface', baseModel({
      chartType: 'surface', categories: ['X1', 'X2'], valAxisMajorUnit: 1,
      series: [series({ values: [1, 2] }), series({ values: [2, 3] })],
    })],
    ['waterfall', baseModel({
      chartType: 'waterfall', categories: ['A', 'B'], series: [series({ values: [2, -1] })],
    })],
    ['funnel', baseModel({
      chartType: 'funnel', categories: ['A', 'B'], series: [series({ values: [2, 1] })],
    })],
    ['3-D column', baseModel({
      chartType: 'clusteredBar', categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
    })],
  ] as const)('paints the shared plot-area frame for %s', (_name, source) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...source,
      plotAreaBg: '123456',
      plotAreaFillPaintAuthored: true,
    }, RECT, 1, 0);
    expect(rec.rects).toContainEqual(expect.objectContaining({ fs: '#123456' }));
  });

  it('renders scatter-series markers and labels over a reversed horizontal category axis', () => {
    const rec = recordingCtx();
    const hiddenAxis = {
      min: 0,
      max: 2,
      title: null,
      hidden: true,
      lineHidden: true,
      majorTickMark: 'none',
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Top', 'Bottom'],
      catAxisOrientation: 'maxMin',
      valMax: 1.4,
      secondaryCatAxis: { ...hiddenAxis, max: 1.4 },
      secondaryValAxis: hiddenAxis,
      series: [
        series({ seriesType: 'bar', values: [0, 0] }),
        series({
          seriesType: 'scatter',
          categories: ['0.2', '1.2'],
          values: [2, 1],
          markerSymbol: 'circle',
          showMarker: true,
          catFormatCode: '0%',
          seriesDataLabels: {
            showCatName: true,
            showSerName: false,
            showVal: false,
            showPercent: false,
          },
        }),
      ],
    }), RECT, 1);

    const top = rec.texts.find(text => text.text === 'Top');
    const bottom = rec.texts.find(text => text.text === 'Bottom');
    expect(top?.y).toBeLessThan(bottom?.y ?? 0);
    const left = rec.texts.find(text => text.text === '20%');
    const right = rec.texts.find(text => text.text === '120%');
    expect(left).toBeDefined();
    expect(right).toBeDefined();
    expect(left?.x).toBeLessThan(right?.x ?? 0);
  });

  it('maps a bar/scatter overlay through its independent authored X/Y axes', () => {
    const rec = recordingCtx();
    const axis = {
      title: null,
      hidden: true,
      lineHidden: true,
      majorTickMark: 'none',
      minorTickMark: 'in',
    } as const;
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [
        series({ seriesType: 'bar', values: [0, 0] }),
        series({
          seriesType: 'scatter', categories: ['0.5', '1.5'], values: [2, 8],
          markerSymbol: 'circle', markerFill: '1696D2', showMarker: true,
          useSecondaryAxis: true,
        }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { grouping: 'clustered', barDirection: 'bar' }),
        plotGroup('scatter', 1, 1, {
          categoryAxis: 'secondary', valueAxis: 'secondary', scatterStyle: 'marker',
        }),
      ],
      secondaryCatAxis: {
        ...axis, min: 0, max: 2, majorUnit: 0.25, minorUnit: 0.05,
      },
      secondaryValAxis: {
        ...axis, min: 0, max: 10, majorUnit: 2, minorUnit: 0.5,
      },
    }), RECT, 1);

    const points = rec.arcs.filter(point => Number.isFinite(point.x) && Number.isFinite(point.y));
    expect(points).toHaveLength(2);
    expect(points[0].x).toBeLessThan(points[1].x);
    expect(points[0].y).toBeGreaterThan(points[1].y);
  });

  it('retains an Office-authored horizontal bar/scatter overlay with ambiguous bottom axes', () => {
    const rec = recordingCtx();
    const hiddenAxis = {
      min: 0,
      max: 2,
      title: null,
      hidden: true,
      lineHidden: true,
      majorTickMark: 'none',
    } as const;
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [
        series({ seriesType: 'bar', values: [0, 0] }),
        series({
          seriesType: 'scatter', categories: ['0.5', '1.5'], values: [2, 1],
          markerSymbol: 'circle', markerFill: '1696D2', showMarker: true,
          useSecondaryAxis: true,
        }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, {
          grouping: 'clustered', barDirection: 'bar', valueAxis: 'unresolved',
        }),
        plotGroup('scatter', 1, 1, {
          categoryAxis: 'unresolved', valueAxis: 'secondary', scatterStyle: 'marker',
        }),
      ],
      secondaryCatAxis: hiddenAxis,
      secondaryValAxis: hiddenAxis,
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('Unsupported chart');
    expect(rec.arcs).toHaveLength(2);
  });

  it('retains a primary stacked-area group with a primary line overlay', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      series: [
        series({ color: '99CCFF', seriesType: 'area', values: [20, 30] }),
        series({ color: '4472C4', seriesType: 'area', values: [10, 15] }),
        series({ color: '000000', seriesType: 'line', values: [25, 35] }),
      ],
      plotGroups: [
        plotGroup('area', 0, 2, { grouping: 'stacked' }),
        plotGroup('line', 2, 1, { grouping: 'standard' }),
      ],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('Unsupported chart');
    expect(rec.filledPaths.length).toBeGreaterThan(0);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#000000')).toBe(true);
  });

  it('retains two same-direction column groups with a secondary line group', () => {
    const rec = recordingCtx();
    const secondary = {
      min: 0,
      max: 1,
      title: null,
      hidden: false,
      lineHidden: false,
      majorTickMark: 'none',
    } as const;
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      secondaryValAxis: secondary,
      series: [
        series({ color: '4472C4', seriesType: 'bar', values: [100, 120] }),
        series({ color: 'EEEEEE', seriesType: 'bar', values: [0.2, 0.4], useSecondaryAxis: true }),
        series({ color: 'ED7D31', seriesType: 'line', values: [0.3, 0.5], useSecondaryAxis: true }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { grouping: 'clustered', barDirection: 'col' }),
        plotGroup('bar', 1, 1, {
          grouping: 'clustered', barDirection: 'col', valueAxis: 'secondary',
        }),
        plotGroup('line', 2, 1, { grouping: 'standard', valueAxis: 'secondary' }),
      ],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('Unsupported chart');
    expect(rec.rects.some(rect => rect.fs === '#4472C4')).toBe(true);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#ED7D31')).toBe(true);
  });

  it('renders a secondary bar group against its right value axis and top category axis', () => {
    const rec = recordingCtx();
    const axisBase = {
      min: null,
      max: null,
      title: null,
      hidden: false,
      lineHidden: false,
      majorTickMark: 'out',
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Primary A'],
      series: [
        series({ color: 'FF0000', values: [80] }),
        series({ color: '00FF00', values: [100] }),
        series({
          color: '0000FF', values: [3], categories: ['Secondary A'],
          useSecondaryAxis: true,
        }),
      ],
      secondaryValAxis: { ...axisBase, min: 0, max: 8 },
      secondaryCatAxis: {
        ...axisBase,
        crosses: 'max',
        title: 'Top Categories',
        fontColor: '123456',
        fontSizeHpt: 800,
        lineColor: '654321',
        lineWidthEmu: 12700,
        tickLabelSkip: 1,
        tickMarkSkip: 1,
      },
    }), RECT, 1);

    const blueBar = rec.rects.find(rect => rect.fs === '#0000FF');
    const redBar = rec.rects.find(rect => rect.fs === '#FF0000');
    expect(blueBar).toBeDefined();
    expect(redBar).toBeDefined();
    // The top category axis crosses its paired value axis at max, so the
    // secondary column starts at the top rule and grows down toward value 3.
    expect(blueBar?.y).toBeLessThan(redBar?.y ?? Number.POSITIVE_INFINITY);
    expect(blueBar?.h).toBeGreaterThan(RECT.h * 0.35);
    expect(rec.texts.find(text => text.text === 'Secondary A')?.fillStyle).toBe('#123456');
    expect(rec.texts.some(text => text.text === 'Top Categories')).toBe(true);
  });

  it.each(['clusteredBar', 'clusteredBarH'] as const)(
    '%s: custom rich bar labels paint bounded inline runs with theme faces',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        themeMajorFontLatin: 'Major Theme',
        themeMinorFontLatin: 'Minor Theme',
        series: [series({
          values: [10],
          dataLabelOverrides: [{
            idx: 0,
            text: 'Major Minor',
            position: 'ctr',
            richRuns: [
              { text: 'Major', fontFace: '+mj-lt', color: '112233' },
              { text: ' Minor', fontFace: '+mn-lt', color: '445566' },
            ],
          }],
        })],
      }), RECT, 1);

      expect(rec.texts.find(call => call.text === 'Major'))
        .toMatchObject({ fillStyle: '#112233' });
      expect(rec.texts.find(call => call.text === 'Major')?.font)
        .toContain('"Major Theme"');
      expect(rec.texts.find(call => call.text === ' Minor'))
        .toMatchObject({ fillStyle: '#445566' });
      expect(rec.texts.find(call => call.text === ' Minor')?.font)
        .toContain('"Minor Theme"');
    },
  );

  it('bar: richRuns do not replace a label composed from show/format flags', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: true,
          formatCode: '0.0%',
        },
        dataLabelOverrides: [{ idx: 0, text: '', richRuns: [{ text: 'stale custom' }] }],
      })],
    }), RECT, 1);

    expect(rec.texts.map(call => call.text)).toContain('100.0%');
    expect(rec.texts.map(call => call.text)).not.toContain('stale custom');
  });

  it('measures the automatic horizontal category-label gutter instead of eliding long labels', () => {
    const rec = recordingCtx();
    const category = 'San Francisco County, California';
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: [category],
      catAxisFontSizeHpt: 1200,
      catAxisFontFace: 'Lato',
      series: [series({ values: [1] })],
    }), RECT, 1);

    expect(rec.texts.some((text) => text.text === category)).toBe(true);
    expect(rec.texts.some((text) => text.text.endsWith('…'))).toBe(false);
  });

  it('does not reserve a horizontal category-label gutter when tick labels are hidden', () => {
    const render = (tickLabelPos: ChartModel['catAxisTickLabelPos']) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBarH',
        categories: ['A category label long enough to affect the gutter'],
        catAxisTickLabelPos: tickLabelPos,
        plotAreaBg: 'ABCDEF',
        series: [series({ values: [1] })],
      }), RECT, 1);
      return rec.rects.find(rect => rect.fs === '#ABCDEF');
    };

    const visible = render('nextTo');
    const hidden = render('none');
    expect(hidden?.x).toBeLessThan(visible?.x ?? 0);
    expect(hidden?.w).toBeGreaterThan(visible?.w ?? Number.POSITIVE_INFINITY);
  });

  it('measures an unauthored horizontal category font at the painted slot size', () => {
    const rect: ChartRect = { x: 0, y: 0, w: 640, h: 720 };
    const render = (fontSizeHpt: number | null) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBarH',
        categories: ['A category label whose width is sensitive to the font size'],
        catAxisFontSizeHpt: fontSizeHpt,
        plotAreaBg: 'ABCDEF',
        series: [series({ values: [1] })],
      }), rect, 1);
      return rec.rects.find(item => item.fs === '#ABCDEF');
    };

    const automatic = render(null);
    const authoredElevenPx = render(1100);
    expect(automatic?.x).toBeCloseTo(authoredElevenPx?.x ?? 0, 6);
    expect(automatic?.w).toBeCloseTo(authoredElevenPx?.w ?? 0, 6);
  });
});

describe('CH1 — negative bar/column values extend from the zero line', () => {
  it('a column chart draws negative bars below the zero line and positive bars above', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    // Two bars, one per category.
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Symmetric data (+10 / -10) → the zero line sits mid-plot and the two bars
    // have equal height. The positive bar's bottom edge equals the negative
    // bar's top edge: they meet at the shared zero line.
    const posBottom = pos.y + pos.h;
    const negTop = neg.y;
    expect(negTop).toBeCloseTo(posBottom, 4); // shared zero line
    // Negative bar hangs BELOW the zero line, positive bar sits ABOVE it.
    expect(neg.y).toBeGreaterThan(pos.y);
    expect(neg.h).toBeGreaterThan(0);
    // Equal magnitudes → equal bar heights.
    expect(neg.h).toBeCloseTo(pos.h, 4);
  });

  it('the value axis includes negative tick labels when data dips below zero', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({ name: 'S', values: [-40] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels.some(l => l.startsWith('-'))).toBe(true);
  });

  it('a horizontal bar chart draws negative bars left of the zero line', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBarH',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, -10] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [pos, neg] = bars;
    // Positive bar starts at the zero line and extends right; negative bar ends
    // at the zero line and extends left, so its right edge equals the positive
    // bar's left edge.
    expect(neg.x + neg.w).toBeCloseTo(pos.x, 4);
    expect(neg.x).toBeLessThan(pos.x);
    expect(neg.w).toBeCloseTo(pos.w, 4);
  });

  it('a deleted value axis uses the default automatic scale instead of visible-tick density', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [20] }),
        series({ name: 'S2', values: [77] }),
      ],
      valAxisHidden: true,
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        x: 0.1,
        y: 0.1,
        w: 0.8,
        h: 0.8,
      },
    }), RECT, 1);

    // With no visible value-axis ticks, Office uses the default five-interval
    // auto-scale target: symmetric padding plus the ceiling 1/2/5 ladder gives
    // data max 97 → major unit 50 → axis max 150.
    const bars = rec.rects;
    expect(bars).toHaveLength(2);
    const totalLength = bars[0].w + bars[1].w;
    expect(totalLength).toBeCloseTo(RECT.w * 0.8 * (97 / 120), 4);
  });

  it('positive-only data keeps the axis anchored at 0 (pre-fix behavior)', () => {
    // Regression guard: min degenerates to 0 so nothing about a positive-only
    // chart changes. Zero-line bottom edge == plot bottom.
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 20] })],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    // All bars share the same bottom edge (the axis at 0), none extend below it.
    const bottoms = bars.map(b => b.y + b.h);
    expect(bottoms[0]).toBeCloseTo(bottoms[1], 4);
    // No negative tick labels.
    expect(rec.texts.every(t => !t.text.startsWith('-'))).toBe(true);
  });

  it('stacked columns accumulate positives up and negatives down separately', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'stackedBar',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),
        series({ name: 'N', values: [-20] }),
      ],
    });
    renderChart(rec.ctx, model, RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line; negative bar below. They meet at
    // the zero line (positive bottom == negative top).
    expect(nBar.y).toBeCloseTo(p.y + p.h, 4);
    expect(nBar.h).toBeGreaterThan(0);
  });
});

describe('CH6 — negative bar data-label placement mirrors the positive convention (§21.2.2.16)', () => {
  // Coverage for drawBarDataLabel's `negative` branch. A single chart holds two
  // categories with a symmetric +37 / -37 value, so BOTH bars share one plot and
  // one axis (a symmetric ±37 range) — the geometry is a clean mirror across the
  // zero line. For each dLblPos the negative label must land on the mirror side
  // of the positive label relative to that shared zero line. "37" / "-37" are
  // not round gridline values, so each data-label text is unambiguous among the
  // recorded fillText calls, and each bar is matched to its label by the shared
  // cross-axis center.
  function renderMirrorBars(
    chartType: 'clusteredBar' | 'clusteredBarH',
    dataLabelPosition: string,
  ): Recorded {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['P', 'N'],
      series: [series({ name: 'S', values: [37, -37] })],
      showDataLabels: true,
      dataLabelPosition,
    }), RECT, 1);
    return rec;
  }
  const labelPos = (rec: Recorded, text: string): TextCall => {
    const hit = rec.texts.find(t => t.text === text);
    expect(hit, `data label "${text}" was drawn`).toBeDefined();
    return hit as TextCall;
  };
  // Match each value bar to its label by the cross-axis center they share
  // (x-center for columns, y-center for horizontal bars).
  const barFor = (rec: Recorded, lbl: TextCall, axis: 'v' | 'h'): RectCall => {
    const center = (b: RectCall) => axis === 'v' ? b.x + b.w / 2 : b.y + b.h / 2;
    const key = axis === 'v' ? lbl.x : lbl.y;
    let best: RectCall | undefined;
    let bestD = Infinity;
    for (const b of rec.rects) {
      const d = Math.abs(center(b) - key);
      if (d < bestD) { bestD = d; best = b; }
    }
    expect(best).toBeDefined();
    return best as RectCall;
  };

  it('honors an explicit non-bold series data-label run property', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        name: 'S',
        values: [37],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          fontBold: false,
          fontSizeHpt: 1100,
          position: 'outEnd',
        },
      })],
      showDataLabels: true,
    }), RECT, 1);
    const label = labelPos(rec, '37');
    expect(label.font).toBe('11px sans-serif');
  });

  describe('vertical columns', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBar', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'v');   // sits ABOVE the zero line
        const negBar = barFor(rec, negLbl, 'v');   // hangs BELOW the zero line
        // Each label is horizontally centered on its own bar.
        expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
        expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
        // Symmetric ±37 → equal bar heights, bars meeting at the shared zero line.
        expect(negBar.h).toBeCloseTo(posBar.h, 3);
        const zeroLine = posBar.y + posBar.h;            // positive bottom == neg top
        expect(negBar.y).toBeCloseTo(zeroLine, 3);
        // The positive bar's value edge is its TOP; the negative's is its BOTTOM.
        const posValueEdge = posBar.y;                   // top edge
        const negValueEdge = negBar.y + negBar.h;        // bottom edge
        if (pos === 'ctr') {
          expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
          expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
          // The two centers are mirror images across the zero line.
          expect(negLbl.y - zeroLine).toBeCloseTo(zeroLine - posLbl.y, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive label offset from its top edge mirrors the negative label
          // offset from its bottom edge (positive sits above → −, negative below → +).
          const posOff = posLbl.y - posValueEdge;
          const negOff = negLbl.y - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: anchored at the zero-line (base) edge for both signs.
          const posBaseEdge = posBar.y + posBar.h;       // bottom (zero line)
          const negBaseEdge = negBar.y;                  // top (zero line)
          const posOff = posLbl.y - posBaseEdge;
          const negOff = negLbl.y - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });

  describe('horizontal bars', () => {
    for (const pos of ['outEnd', 'inEnd', 'inBase', 'ctr']) {
      it(`${pos}: the negative label mirrors the positive label across the zero line`, () => {
        const rec = renderMirrorBars('clusteredBarH', pos);
        const posLbl = labelPos(rec, '37');
        const negLbl = labelPos(rec, '-37');
        const posBar = barFor(rec, posLbl, 'h');   // extends RIGHT of the zero line
        const negBar = barFor(rec, negLbl, 'h');   // extends LEFT of the zero line
        // Each label is vertically centered on its own bar. The recorded rect is
        // fillRect(bx, by, barL, barW), so its HEIGHT is the bar thickness.
        expect(posLbl.y).toBeCloseTo(posBar.y + posBar.h / 2, 4);
        expect(negLbl.y).toBeCloseTo(negBar.y + negBar.h / 2, 4);
        // Symmetric ±37 → equal bar lengths, meeting at the shared zero line.
        expect(negBar.w).toBeCloseTo(posBar.w, 3);
        const zeroLine = posBar.x;                        // positive left == neg right
        expect(negBar.x + negBar.w).toBeCloseTo(zeroLine, 3);
        if (pos === 'ctr') {
          expect(posLbl.x).toBeCloseTo(posBar.x + posBar.w / 2, 4);
          expect(negLbl.x).toBeCloseTo(negBar.x + negBar.w / 2, 4);
          expect(negLbl.x - zeroLine).toBeCloseTo(zeroLine - posLbl.x, 3);
        } else if (pos === 'outEnd' || pos === 'inEnd') {
          // Positive value edge is the RIGHT edge; negative value edge the LEFT.
          const posValueEdge = posBar.x + posBar.w;
          const negValueEdge = negBar.x;
          const posOff = posLbl.x - posValueEdge;
          const negOff = negLbl.x - negValueEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        } else {
          // inBase: zero-line edge. Positive base is the LEFT edge, negative base
          // the RIGHT edge — mirrored across the zero line.
          const posBaseEdge = posBar.x;                  // left (zero line)
          const negBaseEdge = negBar.x + negBar.w;       // right (zero line)
          const posOff = posLbl.x - posBaseEdge;
          const negOff = negLbl.x - negBaseEdge;
          expect(negOff).toBeCloseTo(-posOff, 3);
        }
      });
    }
  });
});

describe('bar point styles, clustered order, and stacked labels', () => {
  it('honors an explicit dPt fill even when varyColors is false', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Overall', 'Other'],
      series: [series({
        color: '1696D2',
        values: [8, 7],
        dataPointColors: ['000000', null],
      })],
      varyColors: false,
    }), RECT, 1);

    expect(rec.rects.map(rect => rect.fs.toUpperCase())).toEqual(['#000000', '#1696D2']);
  });

  it('places series order zero above later series in a horizontal clustered bar', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Black'],
      series: [
        series({ name: 'Trust fund depleted', color: '1696D2', values: [16.2] }),
        series({ name: 'Scheduled benefits', color: '000000', values: [9.9] }),
      ],
    }), RECT, 1);

    const topToBottom = [...rec.rects].sort((a, b) => a.y - b.y);
    expect(topToBottom.map(rect => rect.fs.toUpperCase())).toEqual(['#1696D2', '#000000']);
  });

  it('uses series dLbls and centers their values inside a stacked bar by default', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['Northeast'],
      series: [
        series({
          color: '1696D2',
          values: [0.4],
          valFormatCode: '0.0%',
          seriesDataLabels: {
            showVal: true,
            showCatName: false,
            showSerName: false,
            showPercent: false,
            fontColor: 'FFFFFF',
            fontFace: 'Meiryo UI',
          },
        }),
        series({ color: '000000', values: [0.6] }),
      ],
      // A series-local dLbls block remains operative when the chart-group
      // default is false.
      showDataLabels: false,
      valMax: 1,
    }), RECT, 1);

    const label = rec.texts.find(text => text.text === '40.0%');
    expect(label).toBeDefined();
    const firstSegment = rec.rects[0];
    expect(label?.x).toBeCloseTo(firstSegment.x + firstSegment.w / 2);
    expect(label?.y).toBeCloseTo(firstSegment.y + firstSegment.h / 2);
    expect(label?.align).toBe('center');
    expect(label?.baseline).toBe('middle');
    expect(label?.font).toContain('"Meiryo UI"');
  });

  it('clips stacked geometry to an explicit value-axis maximum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarH',
      categories: ['Exactly 100%', 'Rounded 101%'],
      series: [
        series({ values: [0.5, 0.5] }),
        series({ values: [0.5, 0.51] }),
      ],
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const rows = new Map<number, RectCall[]>();
    for (const rect of rec.rects) {
      const key = Math.round(rect.y * 1000);
      rows.set(key, [...(rows.get(key) ?? []), rect]);
    }
    const widths = [...rows.values()].map(row =>
      Math.max(...row.map(rect => rect.x + rect.w)) - Math.min(...row.map(rect => rect.x)),
    );
    expect(widths).toHaveLength(2);
    expect(widths[0]).toBeCloseTo(widths[1], 6);
  });
});

describe('CH7 — percentStacked normalizes signed values against per-category Σ|v| (§21.2.2.76)', () => {
  // Positive contributions stack up/right, negatives down/left; each series is
  // normalized to (v / Σ|v|)·100 so the axis spans −100..100.
  it.each([
    'stackedBarPct',
    'stackedLinePct',
    'stackedAreaPct',
  ] as const)('%s scales fractional OOXML axis units to percentage points', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [25, 75] }),
        series({ name: 'S2', values: [75, 25] }),
      ],
      // The chart stores percent-axis values as ratios. The renderer's stacked
      // geometry uses percentage points internally, so 0.5 must become the
      // 50-point interval Excel displays as 50%, not a 0.5-point interval.
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
    }), RECT, 1);

    const labels = rec.texts
      .map(t => t.text)
      .filter(t => /^-?\d+(?:\.\d+)?%$/.test(t));
    expect(labels).toEqual(['0%', '50%', '100%']);
  });

  it('column percent-axis labels honor the font size declared in valAx txPr', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
      // DrawingML run sizes are hundredths of a point: 1100 = 11 pt.
      valAxisFontSizeHpt: 1100,
    }), RECT, 4 / 3);

    const percentLabels = rec.texts.filter(t => /^(?:0|50|100)%$/.test(t.text));
    expect(percentLabels).toHaveLength(3);
    for (const label of percentLabels) {
      const fontPx = Number(/^([\d.]+)px/.exec(label.font ?? '')?.[1]);
      expect(fontPx).toBeCloseTo(11 * 4 / 3, 5);
    }
  });

  it('column category-axis labels honor the font size inherited from chartSpace txPr', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['T1', 'T2'],
      series: [series({ name: 'S1', values: [10, 20] })],
      // The shared parser resolves chartSpace/txPr onto both axis fields when
      // the individual axes have no txPr. 1800 = 18 pt.
      catAxisFontSizeHpt: 1800,
    }), RECT, 4 / 3);

    const categoryLabels = rec.texts.filter(t => /^T[12]$/.test(t.text));
    expect(categoryLabels).toHaveLength(2);
    for (const label of categoryLabels) {
      const fontPx = Number(/^(?:bold )?([\d.]+)px/.exec(label.font ?? '')?.[1]);
      expect(fontPx).toBeCloseTo(18 * 4 / 3, 5);
    }
  });

  it('keeps explicit-size value-axis labels inside a correctly authored inner manual-layout frame', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
      valAxisFontSizeHpt: 1100,
      // ECMA-376 §21.2.2.89: an inner target describes the data region,
      // excluding axes and labels. The producer therefore reserves the label
      // gutter in the authored x offset rather than relying on auto-layout.
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        x: 0.184,
        y: 0.046,
        w: 0.728,
        h: 0.784,
      },
    }), RECT, 4 / 3);

    const label = rec.texts.find(t => t.text === '100%');
    expect(label).toBeDefined();
    expect(label!.align).toBe('right');
    expect(label!.x - (label!.width ?? 0)).toBeGreaterThanOrEqual(RECT.x + 4);
  });

  it('scales explicit fractional percent-axis bounds before plotting', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'S1', values: [25] }),
        series({ name: 'S2', values: [75] }),
      ],
      valMin: 0,
      valMax: 1,
      valAxisMajorUnit: 0.5,
      valAxisFormatCode: '0%',
    }), RECT, 1);

    const labels = rec.texts
      .map(t => t.text)
      .filter(t => /^-?\d+(?:\.\d+)?%$/.test(t));
    expect(labels).toEqual(['0%', '50%', '100%']);
  });

  it('vertical percentStacked: positives stack above zero, negatives below, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30
        series({ name: 'N', values: [-10] }),  // -10  → Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar sits above the zero line, negative bar below; they meet at it.
    expect(nBar.y).toBeCloseTo(p.y + p.h, 3);          // shared zero line
    expect(nBar.y).toBeGreaterThan(p.y);               // negative is lower
    // Normalized magnitudes: +30/40 = 75% up, -10/40 = 25% down. Same axis
    // scale (px per percent) → the positive bar is 3× the negative bar's height.
    expect(p.h / nBar.h).toBeCloseTo(3, 2);
    // The value axis carries the ±100 percentStacked gridlines (plus headroom,
    // so the outermost ticks sit at ±120, matching the line/area pct convention).
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('horizontal percentStacked: positives stack right, negatives left, normalized to Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarHPct',
      categories: ['A'],
      series: [
        series({ name: 'P', values: [30] }),   // +30 → right
        series({ name: 'N', values: [-10] }),  // -10 → left, Σ|v| = 40
      ],
    }), RECT, 1);
    const bars = rec.rects;
    expect(bars.length).toBe(2);
    const [p, nBar] = bars;
    // Positive bar extends right of the zero line, negative left; they meet at it.
    expect(nBar.x + nBar.w).toBeCloseTo(p.x, 3);       // shared zero line
    expect(nBar.x).toBeLessThan(p.x);                  // negative is to the left
    // +30/40 = 75% right vs -10/40 = 25% left → 3× the width.
    expect(p.w / nBar.w).toBeCloseTo(3, 2);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    expect(nums).toContain(100);
    expect(nums).toContain(-100);
    expect(Math.min(...nums)).toBeLessThanOrEqual(-100);
    expect(Math.min(...nums)).toBeGreaterThanOrEqual(-120);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('multi-category percentStacked: each category normalizes to its own Σ|v|', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'P', values: [10, 40] }),  // A: Σ|v|=20  B: Σ|v|=50
        series({ name: 'N', values: [-10, -10] }),
      ],
    }), RECT, 1);
    const bars = rec.rects;
    // Two categories × two series = four bars, in draw order: A/P, A/N, B/P, B/N.
    expect(bars.length).toBe(4);
    const [aP, aN, bP, bN] = bars;
    // Category A: 10 and -10 of Σ|v|=20 → 50% up, 50% down → equal heights.
    expect(aP.h).toBeCloseTo(aN.h, 2);
    // Category B: 40 and -10 of Σ|v|=50 → 80% up, 20% down → positive is 4× taller.
    expect(bP.h / bN.h).toBeCloseTo(4, 2);
    // Per-category normalization (not a global Σ): A's +50% bar and B's +80% bar
    // are NOT the same height even though A/P is the larger raw share of A.
    expect(bP.h).toBeGreaterThan(aP.h);
  });
});

describe('CH2 — stackedLine / stackedLinePct stack cumulatively', () => {
  it('draws every authored non-empty sparse category label when tickLblSkip is absent', () => {
    const rec = recordingCtx();
    const categories = Array.from({ length: 25 }, () => '');
    const expected: string[] = [];
    for (let index = 1, year = 2000; year <= 2022; index += 2, year += 2) {
      categories[index] = String(year);
      expected.push(String(year));
    }
    // The final source row also carries 2022. Excel paints both adjacent labels
    // rather than letting an auto-collision heuristic discard the authored
    // sparse sequence.
    categories[24] = '2022';

    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories,
      series: [series({ values: categories.map((_, index) => index + 1) })],
      catAxisFontSizeHpt: 1200,
    }), RECT, 1);

    const yearLabels = rec.texts
      .map(text => text.text)
      .filter(text => /^20\d{2}$/.test(text));
    expect(yearLabels).toEqual([...expected, '2022']);
  });

  it('stackedLine plots the second series at the cumulative sum', () => {
    // Two flat series (all 10, all 20). Stacked, the second line rides at
    // y=30 across every category; unstacked it would ride at y=20. We detect
    // stacking by the axis maximum: a cumulative 30 forces a taller axis than
    // an un-stacked max of 20 would.
    const stackedRec = recordingCtx();
    renderChart(stackedRec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const plainRec = recordingCtx();
    renderChart(plainRec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'S1', values: [10, 10, 10] }),
        series({ name: 'S2', values: [20, 20, 20] }),
      ],
    }), RECT, 1);

    const stackedTop = Math.max(...stackedRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    const plainTop = Math.max(...plainRec.texts
      .map(t => Number(t.text)).filter(v => Number.isFinite(v)));
    // Stacking pushes the cumulative maximum (30) above the plain per-series
    // maximum (20), so the auto axis top must be strictly higher.
    expect(stackedTop).toBeGreaterThan(plainTop);
  });

  it('stackedLinePct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLinePct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums.
    expect(nums).toContain(100);
    // ...and the axis top is a round value just above 100 (headroom), never the
    // raw cumulative magnitude of 40.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('plain line is unaffected (per-series max drives the axis)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 10] }),
        series({ name: 'S2', values: [20, 20] }),
      ],
    }), RECT, 1);
    const top = Math.max(...rec.texts.map(t => Number(t.text)).filter(Number.isFinite));
    // Un-stacked: axis reflects the single-series max (20) plus headroom, not 30.
    expect(top).toBeLessThan(30);
  });
});

describe('CH4 — stackedAreaPct normalizes like the line/bar percentStacked convention', () => {
  it('keeps value labels inside an authored outer plot-area rectangle', () => {
    const rec = recordingCtx();
    const outerX = 0.007891414141414141 * 760;
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['2016', '2017'],
      series: [
        series({ values: [0.45, 0.4] }),
        series({ values: [0.55, 0.6] }),
      ],
      valMax: 1,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisFontSizeHpt: 1200,
      catAxisFontSizeHpt: 1200,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge',
        x: 0.007891414141414141,
        y: 0.1949702068511199,
        w: 0.9732744107744108,
        h: 0.6791097091286356,
      },
    }), { x: 0, y: 0, w: 760, h: 560 }, 1);

    const topTick = rec.texts.find(text => text.text === '1.0');
    expect(topTick).toBeDefined();
    expect(topTick?.align).toBe('right');
    const tickLeft = (topTick?.x ?? 0) - (topTick?.width ?? 0);
    // CT_LayoutTarget defaults to `outer`: its left edge includes the value
    // labels, while the inner data rectangle begins after their measured width
    // and authored-font gap. The label must not be pushed outside chart space.
    expect(tickLeft).toBeCloseTo(outerX + 1.5, 5);
    const plotBg = rec.rects.find(rect => rect.fs === '#ABCDEF');
    expect(plotBg?.x).toBeGreaterThan(outerX);
  });

  it('honors the authored inner plot-area rectangle for area charts', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'],
      series: [series({ values: [1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1] })],
      plotAreaBg: 'ABCDEF',
      catAxisTickLabelSkip: 5,
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', layoutTarget: 'inner',
        x: 0.2, y: 0.25, w: 0.5, h: 0.4,
      },
    }), RECT, 1);
    expect(rec.rects).toContainEqual({
      x: RECT.w * 0.2,
      y: RECT.h * 0.25,
      w: RECT.w * 0.5,
      h: RECT.h * 0.4,
      fs: '#ABCDEF',
    });
    expect(rec.texts.some(text => text.text === 'A')).toBe(true);
    expect(rec.texts.some(text => text.text === 'F')).toBe(true);
    expect(rec.texts.some(text => text.text === 'K')).toBe(true);
    expect(rec.texts.some(text => text.text === 'B')).toBe(false);
  });

  it('stackedAreaPct normalizes each category to 100%', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedAreaPct',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // The cumulative top series always reaches exactly 100% per category, so the
    // axis carries a 100 gridline. Raw magnitudes (max cumulative 40) never
    // appear — the axis is normalized, not driven by the raw sums (this was Red
    // before the fix: stackedAreaPct was treated identically to stackedArea, so
    // the axis topped out at the raw cumulative 40 instead of 100).
    expect(nums).toContain(100);
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(100);
    expect(Math.max(...nums)).toBeLessThanOrEqual(120);
  });

  it('stackedArea (non-percent) is unaffected — axis reflects the raw cumulative sum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A', 'B'],
      series: [
        series({ name: 'S1', values: [10, 30] }),
        series({ name: 'S2', values: [30, 10] }),
      ],
    }), RECT, 1);
    const nums = rec.texts.map(t => Number(String(t.text).replace('%', '')))
      .filter(v => Number.isFinite(v));
    // Raw cumulative max per category is 40 (10+30 / 30+10); the axis must scale
    // to that magnitude, not be normalized to 100.
    expect(Math.max(...nums)).toBeGreaterThanOrEqual(40);
    expect(nums).not.toContain(100);
  });
});

describe('axis display units', () => {
  it('divides scatter X/Y tick text and paints each authored display-unit label', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: [],
      series: [series({
        values: [200, 800],
        categories: ['1000', '4000'],
        showMarker: true,
      })],
      catAxisMin: 0,
      catAxisMax: 4_000,
      catAxisMajorUnit: 1_000,
      valMin: 0,
      valMax: 800,
      valAxisMajorUnit: 200,
      catAxisDisplayUnits: {
        divisor: 1_000,
        builtInUnit: 'thousands',
        label: { text: 'X thousands', manualLayout: { x: 0.7, y: 0.8 } },
      },
      valAxisDisplayUnits: {
        divisor: 100,
        builtInUnit: 'hundreds',
        label: { manualLayout: { x: 0.1, y: 0.1 }, fontBold: true },
      },
    }), { x: 0, y: 0, w: 500, h: 300 }, 1);

    const text = rec.texts.map(item => item.text);
    expect(text).toContain('X thousands');
    expect(text).toContain('Hundreds');
    expect(text).not.toContain('1000');
    expect(text).not.toContain('200');
    expect(text).toEqual(expect.arrayContaining(['1', '2', '4', '8']));
  });

  it('formats cartesian 3-D value-axis ticks with the same display units', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ values: [200, 800] })],
      valMin: 0,
      valMax: 800,
      valAxisMajorUnit: 200,
      valAxisDisplayUnits: {
        divisor: 100,
        builtInUnit: 'hundreds',
        label: null,
      },
      threeD: { rotationX: 15, rotationY: 20 },
    }), { x: 0, y: 0, w: 500, h: 300 }, 1);

    const text = rec.texts.map(item => item.text);
    expect(text).toEqual(expect.arrayContaining(['0', '2', '4', '6', '8']));
    expect(text).not.toContain('200');
    expect(text).not.toContain('800');
  });

  it.each(['scatter', 'line', 'clusteredBar'] as const)(
    '%s applies the associated value-axis display unit to showVal data labels',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: chartType === 'scatter' ? [] : ['A'],
        series: [series({
          values: [8],
          categories: chartType === 'scatter' ? ['1000'] : undefined,
          showMarker: true,
          seriesDataLabels: {
            showVal: true,
            showCatName: false,
            showSerName: false,
            showPercent: false,
            fontColor: 'FF0000',
            formatCode: '0.00',
          },
        })],
        catAxisMin: chartType === 'scatter' ? 0 : undefined,
        catAxisMax: chartType === 'scatter' ? 2_000 : undefined,
        valMin: 0,
        valMax: 100,
        valAxisDisplayUnits: {
          divisor: 100,
          builtInUnit: 'hundreds',
          label: null,
        },
      }), RECT, 1);

      expect(rec.texts.find(text => text.fillStyle === '#FF0000')?.text).toBe('0.08');
    },
  );

  it('applies value-axis display units to chart-group showVal labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({ values: [8], showMarker: true })],
      showDataLabels: true,
      dataLabelFormatCode: '0.00',
      dataLabelFontColor: 'FF0000',
      valMin: 0,
      valMax: 100,
      valAxisDisplayUnits: {
        divisor: 100,
        builtInUnit: 'hundreds',
        label: null,
      },
    }), RECT, 1);

    expect(rec.texts).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: '0.08', fillStyle: '#FF0000' }),
    ]));
  });

  it('uses the secondary scatter Y-axis display unit for that group only', () => {
    const rec = recordingCtx();
    const hiddenAxis = {
      min: 0,
      max: 20,
      title: null,
      hidden: true,
      lineHidden: true,
      majorTickMark: 'none',
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: [],
      series: [
        series({ categories: ['1'], values: [8] }),
        series({
          categories: ['1000'],
          values: [8],
          useSecondaryAxis: true,
          seriesDataLabels: {
            showVal: true,
            showCatName: false,
            showSerName: false,
            showPercent: false,
            fontColor: 'FF0000',
            formatCode: '0.0',
          },
        }),
      ],
      catAxisMin: 0,
      catAxisMax: 2,
      valMin: 0,
      valMax: 10,
      valAxisDisplayUnits: { divisor: 100, builtInUnit: 'hundreds', label: null },
      secondaryCatAxis: { ...hiddenAxis, max: 2_000 },
      secondaryValAxis: {
        ...hiddenAxis,
        displayUnits: { divisor: 10, builtInUnit: null, label: null },
      },
    }), RECT, 1);

    expect(rec.texts.find(text => text.fillStyle === '#FF0000')?.text).toBe('0.8');
  });

  it('applies value-axis display units to cartesian 3-D showVal labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        values: [8],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          fontColor: 'FF0000',
          formatCode: '0.00',
        },
      })],
      valMin: 0,
      valMax: 100,
      valAxisDisplayUnits: {
        divisor: 100,
        builtInUnit: 'hundreds',
        label: null,
      },
      threeD: { rotationX: 15, rotationY: 20 },
    }), RECT, 1);

    expect(rec.texts.find(text => text.fillStyle === '#FF0000')?.text).toBe('0.08');
  });
});

describe('ECMA-376 §21.2.2.89 — omitted layoutTarget defaults to outer', () => {
  const outer = {
    xMode: 'edge' as const,
    yMode: 'edge' as const,
    x: 0.01,
    y: 0.2,
    w: 0.97,
    h: 0.68,
  };

  it('line: keeps the formatted value labels inside the outer rectangle', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['2000', '2002'],
      series: [series({ values: [200, 1_400] })],
      valMin: 0,
      valMax: 1_400,
      valAxisMajorUnit: 200,
      valAxisFormatCode: '"$"#,##0',
      valAxisFontSizeHpt: 1000,
      catAxisFontSizeHpt: 1000,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: outer,
    }), { x: 0, y: 0, w: 700, h: 420 }, 1);

    const topTick = rec.texts.find(text => text.text === '$1,400');
    expect(topTick).toBeDefined();
    const tickLeft = (topTick?.x ?? 0) - (topTick?.width ?? 0);
    expect(tickLeft).toBeCloseTo(8.5, 5);
    expect(rec.rects.find(rect => rect.fs === '#ABCDEF')?.x).toBeGreaterThan(7);
  });

  it('column: removes the category-label band from the outer plot height', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      valAxisHidden: true,
      catAxisFontSizeHpt: 1000,
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: outer,
    }), { x: 0, y: 0, w: 700, h: 420 }, 1);

    const plot = rec.rects.find(rect => rect.fs === '#ABCDEF');
    expect(plot).toBeDefined();
    expect(plot?.x).toBeCloseTo(7, 5);
    expect(plot?.h).toBeLessThan(0.68 * 420);
    expect(rec.texts.find(text => text.text === 'A')?.y).toBeGreaterThan((plot?.y ?? 0) + (plot?.h ?? 0));
  });

  it('uses the same measured axis-label insets for line and area geometry', () => {
    const plots = (chartType: 'line' | 'area') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['2016', '2017'],
        series: [series({ values: [20, 40] })],
        valMax: 40,
        valAxisMajorUnit: 10,
        valAxisFontSizeHpt: 1000,
        catAxisFontSizeHpt: 1000,
        plotAreaBg: 'ABCDEF',
        plotAreaManualLayout: outer,
      }), { x: 0, y: 0, w: 700, h: 420 }, 1);
      return rec.rects.find(rect => rect.fs === '#ABCDEF');
    };

    expect(plots('line')).toEqual(plots('area'));
  });
});

describe('CH5 — category axis numFmt applies to category tick labels (§21.2.2.71)', () => {
  // dateAx / numeric category axes cache the categories as Excel serial numbers
  // ("44927"). Before the fix the renderer drew those raw serials; now the
  // catAxisFormatCode is applied so a time-series line/bar shows real dates.
  const DATE_CATS = ['44927', '44958', '44986']; // 2023-01-01 / 02-01 / 03-01

  it('a line chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels).toContain('2/1/2023');
    expect(labels).toContain('3/1/2023');
    // The raw serials must NOT appear as category labels anymore.
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('a column chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('clips column marks before an authored date-axis minimum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['45600', '45630', '45660'],
      catAxisIsDate: true,
      catAxisBaseTimeUnit: 'days',
      catAxisMajorTimeUnit: 'days',
      catAxisMajorUnit: 30,
      catAxisMin: 45630,
      catAxisMax: 45660,
      series: [series({ color: '4472C4', values: [10, 20, 30] })],
    }), RECT, 1);

    const bars = rec.rects.filter(rect => rect.fs === '#4472C4');
    expect(bars).toHaveLength(2);
    expect(bars.every(bar => bar.x >= 0)).toBe(true);
  });

  it('keeps a zero-gap overlay bar group as one continuous date range', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['45658', '45689', '45717'],
      catAxisIsDate: true,
      catAxisBaseTimeUnit: 'months',
      catAxisMajorTimeUnit: 'months',
      catAxisMajorUnit: 1,
      catAxisMin: 45658,
      series: [
        series({
          color: '4472C4', values: [10, 20, 30], seriesType: 'bar',
          barGroupIndex: 0, barGroupGapWidth: 150,
        }),
        series({
          color: 'D3D3D333', values: [1, 1, 0], seriesType: 'bar', useSecondaryAxis: true,
          barGroupIndex: 1, barGroupGapWidth: 0, barGroupOverlap: 100,
        }),
      ],
      secondaryValAxis: {
        min: 0.06, max: 0.08, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'cross',
      },
    }), RECT, 1);

    const range = rec.rects.filter(rect => rect.fs === '#D3D3D333' && rect.h > 0);
    expect(range).toHaveLength(2);
    expect(range[0].x + range[0].w).toBeCloseTo(range[1].x);
  });

  it('keeps fractional calendar-unit labels fail-closed while sharing series positions', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['45292', '45302', '45383'],
      catAxisIsDate: true,
      catAxisBaseTimeUnit: 'days',
      catAxisMajorTimeUnit: 'months',
      catAxisMajorUnit: 1.5,
      catAxisMin: 45292,
      catAxisMax: 45383,
      catAxisFormatCode: 'm/d/yyyy',
      series: [
        series({ color: '4472C4', lineColor: '4472C4', values: [10, 20, 30], showMarker: false }),
        series({
          color: 'ED7D31', lineColor: 'ED7D31', values: [30, 20, 10], showMarker: false,
          useSecondaryAxis: true,
        }),
      ],
      secondaryValAxis: {
        min: 0, max: 40, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none',
      },
    }), RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels.some(label => label.includes('/2024'))).toBe(false);
    const primary = rec.strokes.find(stroke => stroke.ss === '#4472C4' && stroke.points.length === 3);
    const secondary = rec.strokes.find(stroke => stroke.ss === '#ED7D31' && stroke.points.length === 3);
    expect(primary).toBeDefined();
    expect(secondary).toBeDefined();
    expect(secondary!.points.map(point => point.x))
      .toEqual(primary!.points.map(point => point.x));
    expect(primary!.points[1]!.x - primary!.points[0]!.x)
      .toBeLessThan(primary!.points[2]!.x - primary!.points[1]!.x);
  });

  it('positions horizontal bar clusters within their owning group', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['A'],
      series: [
        series({ color: 'FF0000', values: [10], barGroupIndex: 0, barGroupGapWidth: 100 }),
        series({ color: '0000FF', values: [20], barGroupIndex: 0, barGroupGapWidth: 100 }),
        series({ color: '00FF00', values: [15], barGroupIndex: 1, barGroupGapWidth: 0 }),
      ],
    }), RECT, 1);

    const red = rec.rects.find(rect => rect.fs === '#FF0000')!;
    const blue = rec.rects.find(rect => rect.fs === '#0000FF')!;
    const overlay = rec.rects.find(rect => rect.fs === '#00FF00')!;
    expect(overlay.y).toBeLessThanOrEqual(Math.min(red.y, blue.y));
    expect(overlay.y + overlay.h).toBeGreaterThanOrEqual(Math.max(red.y + red.h, blue.y + blue.h));
  });

  it('keeps stacked and clustered semantics local to each bar group', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [
        series({
          color: 'FF0000', values: [10], barGroupIndex: 0,
          barGroupDirection: 'col', barGroupGrouping: 'stacked',
        }),
        series({
          color: '0000FF', values: [20], barGroupIndex: 0,
          barGroupDirection: 'col', barGroupGrouping: 'stacked',
        }),
        series({
          color: '00FF00', values: [15], barGroupIndex: 1,
          barGroupDirection: 'col', barGroupGrouping: 'clustered',
        }),
      ],
    }), RECT, 1);

    const red = rec.rects.find(rect => rect.fs === '#FF0000')!;
    const blue = rec.rects.find(rect => rect.fs === '#0000FF')!;
    const green = rec.rects.find(rect => rect.fs === '#00FF00')!;
    expect(red.x).toBeCloseTo(blue.x);
    expect(red.w).toBeCloseTo(blue.w);
    expect(Math.min(red.y + red.h, blue.y + blue.h)).toBeCloseTo(
      Math.max(red.y, blue.y),
    );
    expect(green.x).toBeLessThanOrEqual(red.x);
    expect(green.x + green.w).toBeGreaterThanOrEqual(red.x + red.w);
  });

  it('normalizes a secondary-axis percent-stacked bar group before scaling', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      secondaryValAxis: {
        min: null, max: null, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none', formatCode: '0%',
      },
      series: [
        series({
          color: '808080', values: [10], barGroupIndex: 0,
          barGroupDirection: 'col', barGroupGrouping: 'clustered',
        }),
        series({
          color: 'FF0000', values: [1], barGroupIndex: 1,
          barGroupDirection: 'col', barGroupGrouping: 'percentStacked',
          useSecondaryAxis: true,
        }),
        series({
          color: '0000FF', values: [1], barGroupIndex: 1,
          barGroupDirection: 'col', barGroupGrouping: 'percentStacked',
          useSecondaryAxis: true,
        }),
      ],
    }), RECT, 1);

    const red = rec.rects.find(rect => rect.fs === '#FF0000')!;
    const blue = rec.rects.find(rect => rect.fs === '#0000FF')!;
    expect(red.h).toBeGreaterThan(0);
    expect(blue.h).toBeGreaterThan(0);
    expect(Math.min(red.y + red.h, blue.y + blue.h)).toBeCloseTo(
      Math.max(red.y, blue.y),
    );
    expect(rec.texts.map(text => text.text)).toContain('100%');
  });

  it('keeps horizontal date-axis category spacing when majorUnit is automatic', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['1', '100', '101'],
      catAxisIsDate: true,
      catAxisBaseTimeUnit: 'days',
      series: [series({ color: '4472C4', values: [10, 20, 30] })],
    }), RECT, 1);

    const centers = rec.rects
      .filter(rect => rect.fs === '#4472C4')
      .map(rect => rect.y + rect.h / 2);
    expect(centers).toHaveLength(3);
    expect(Math.abs(centers[1]! - centers[0]!))
      .toBeGreaterThan(Math.abs(centers[2]! - centers[1]!) * 50);
  });

  it('a horizontal bar chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('an area chart formats numeric-serial categories through the date code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: DATE_CATS,
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('1/1/2023');
    expect(labels.some(l => l === '44927')).toBe(false);
  });

  it('string categories stay verbatim even when a format code is present', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Q1', 'Q2', 'Q3'],
      catAxisFormatCode: 'm/d/yyyy',
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('Q1');
    expect(labels).toContain('Q2');
    expect(labels).toContain('Q3');
  });

  it('numeric categories with no format code render as raw text (unchanged)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: DATE_CATS,
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('44927');
    expect(labels.some(l => l === '1/1/2023')).toBe(false);
  });
});

describe('CH3 — labels are locale-independent (§18.8.30)', () => {
  // `toLocaleString()` groups thousands in every common locale, so an explicit
  // no-separator format code ("0") is the discriminator: the §18.8.30 engine
  // honors it (no commas), while toLocaleString ignores it and always inserts
  // the host locale's group separator. The old code called toLocaleString and
  // dropped the format code entirely, so these tests were Red before the fix.
  it('waterfall data labels honor the format code (no host-locale grouping)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1234567, 0] })],
      subtotalIndices: [1],
      dataLabelFormatCode: '0',
      showDataLabels: true,
    }), RECT, 1);
    // The 1234567 subtotal bar's label must be un-grouped ("1234567"), proving
    // it went through the §18.8.30 engine rather than toLocaleString().
    expect(rec.texts.some(t => t.text.includes('1234567'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('1,234,567'))).toBe(true);
  });

  it('waterfall negative data labels honor the authored negative number-format section', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Drop', 'End'],
      series: [series({
        name: 'W',
        values: [2, -0.2, 1.8],
        valFormatCode: '_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"??_);_(@_)',
      })],
      subtotalIndices: [2],
      showDataLabels: true,
    }), RECT, 1);
    expect(rec.texts.some(t => t.text.includes('(0.2)'))).toBe(true);
    expect(rec.texts.every(t => !t.text.includes('△'))).toBe(true);
  });

  it('waterfall preserves rich point-run visibility and linked fallback color', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'HIDDENVISIBLE',
          richRuns: [
            { text: 'HIDDEN', colorPaintAuthored: true, colorHidden: true },
            { text: 'VISIBLE' },
          ],
        }],
      })],
      chartStyleRoles: { dataLabel: { fontColor: '008000' } },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === 'HIDDEN')).toBe(false);
    expect(rec.texts).toContainEqual(expect.objectContaining({ text: 'VISIBLE', fillStyle: '#008000' }));
  });

  it('waterfall value-axis labels honor the format code (through the §18.8.30 engine)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'End'],
      series: [series({ name: 'W', values: [1000000, 0] })],
      subtotalIndices: [1],
      valAxisFormatCode: '0',
    }), RECT, 1);
    // A no-separator format code must suppress grouping. The old code ignored
    // valAxisFormatCode and always grouped via toLocaleString(), so a "1,000,000"
    // tick label would appear — after the fix the ticks are un-grouped.
    expect(rec.texts.every(t => !t.text.includes('1,000,000'))).toBe(true);
    expect(rec.texts.some(t => /^\d{4,}$/.test(t.text))).toBe(true);
  });

  it('waterfall renders ChartEx titles, axis fonts, wrapped categories, and themed point roles', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      title: 'EBITDA bridge',
      titleFontSizeHpt: 1400,
      titleFontBold: false,
      titleFontFace: 'Calibri',
      valAxisTitle: '$ in million',
      valAxisTitleFontSizeHpt: 900,
      valAxisTitleFontBold: false,
      valAxisTitleFontFace: 'Calibri',
      valAxisFontSizeHpt: 900,
      valAxisFontFace: 'Calibri',
      catAxisFontSizeHpt: 900,
      catAxisFontFace: 'Calibri',
      dataLabelFontSizeHpt: 900,
      dataLabelFontBold: false,
      dataLabelFontFace: 'Calibri',
      showDataLabels: true,
      categories: [
        'EBITDA FY21',
        'Change in Revenues',
        'Change in Variable costs',
        'Change in Opex',
        'EBITDA FY22',
      ],
      series: [series({ name: 'W', values: [4.2, 0.3, -0.2, 1.0, 5.3] })],
      subtotalIndices: [4],
      barGapWidth: 50,
      chartexAccents: ['E6E7E8', 'F57A16', '1E8496', '000000', '000000', '000000'],
    }), RECT, 1);

    const fills = rec.rects.map(rect => rect.fs.toUpperCase());
    expect(fills).toEqual(['#E6E7E8', '#E6E7E8', '#F57A16', '#E6E7E8', '#1E8496']);
    expect(rec.texts.some(text =>
      text.text === 'EBITDA bridge' &&
      text.font?.includes('14px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text =>
      text.text === '$ in million' &&
      text.font?.includes('9px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text =>
      text.text === '4.2' &&
      text.font?.startsWith('9px') &&
      text.font.includes('Calibri')
    )).toBe(true);
    expect(rec.texts.some(text => text.text.includes('Variable'))).toBe(true);
    expect(rec.texts.some(text => text.text === 'costs')).toBe(true);
  });

  it('uses the ChartEx seriesLine role for waterfall connectors', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({ name: 'W', values: [10, -2, 8] })],
      subtotalIndices: [2],
      chartexDataPointStyle: { lineColors: ['C00000'] },
      chartexDataPointLineStyle: { lineColors: ['70AD47'], lineWidthEmu: 12700 },
      chartexSeriesLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 25400 },
    }), RECT, 1);
    const connectors = rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0');
    expect(connectors).toHaveLength(2);
    expect(connectors.every(segment => segment.lw === 2)).toBe(true);
  });

  it('uses the linked seriesLine color and 0.75pt width for waterfall connectors', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({ name: 'W', values: [10, -2, 8] })],
      subtotalIndices: [2],
      chartexSeriesLineStyle: {
        lineColors: ['D9D9D9'],
        lineWidthEmu: 9525,
        lineCap: 'flat',
        lineJoin: 'round',
      },
    }), RECT, 1);

    const connectors = rec.segs.filter(segment => segment.ss.toLowerCase() === '#d9d9d9');
    expect(connectors).toHaveLength(2);
    expect(connectors.every(segment => segment.lw === 0.75)).toBe(true);
  });

  it('keeps a direct Waterfall connector stroke authoritative over linked NoStyle', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({
        name: 'W', values: [10, -2, 8],
        chartexStyle: {
          lineColors: ['123456'], lineWidthEmu: 25400,
          lineDash: 'dash', lineCap: 'rnd', lineJoin: 'bevel',
        },
      })],
      subtotalIndices: [2],
      chartexSeriesLineStyle: { lineHidden: true, lineNoStyle: true },
    }), RECT, 1);

    const connectors = rec.strokes.filter(segment => segment.ss.toLowerCase() === '#123456');
    expect(connectors).toHaveLength(2);
    expect(connectors.every(segment =>
      segment.lw === 2 && segment.dash.length > 0
      && segment.cap === 'round' && segment.join === 'bevel'
    )).toBe(true);
  });

  it('keeps direct Waterfall connector noFill authoritative over a linked stroke', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({
        name: 'W', values: [10, -2, 8],
        chartexStyle: { lineHidden: true, lineWidthEmu: 25400, lineDash: 'dash' },
      })],
      subtotalIndices: [2],
      chartexSeriesLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 12700 },
    }), RECT, 1);

    expect(rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0'))
      .toHaveLength(0);
  });

  it('uses the Office-observed semantic connector stroke for linked NoStyle', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['1', '2', '3'],
      series: [series({ name: 'W', values: [1, 1, 1] })],
      chartexSeriesLineStyle: { lineHidden: true, lineNoStyle: true },
    }), RECT, 1);

    const connectors = rec.segs.filter(segment => segment.ss.toLowerCase() === '#000000');
    expect(connectors).toHaveLength(2);
    expect(connectors.every(segment => segment.lw === 0.75)).toBe(true);
  });

  it('keeps semantic data-point fills when the ChartEx series shape has noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Drop', 'End'],
      series: [series({
        name: 'W',
        values: [10, -2, 8],
        chartexStyle: { fillHidden: true, lineColors: ['4472C4'] },
      })],
      subtotalIndices: [2],
      chartexDataPointStyle: { fillColors: ['5B9BD5', 'ED7D31', 'A5A5A5'] },
    }), RECT, 1);

    expect(rec.rects.map(rect => rect.fs.toUpperCase())).toEqual([
      '#5B9BD5', '#ED7D31', '#A5A5A5',
    ]);
  });

  it('suppresses waterfall connectors when CT_SeriesElementVisibilities says false', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({ name: 'W', values: [10, -2, 8] })],
      subtotalIndices: [2],
      chartexConnectorLines: false,
      chartexSeriesLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 25400 },
    }), RECT, 1);
    expect(rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0')).toHaveLength(0);
  });

  it('does not invent waterfall value labels when no data-label definition exists', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Change', 'End'],
      series: [series({ name: 'W', values: [10, -2, 8] })],
      subtotalIndices: [2],
      valAxisHidden: true,
      catAxisHidden: true,
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).not.toEqual(
      expect.arrayContaining(['10', '-2', '8']),
    );
  });

  it('omits automatic value-axis labels for an all-increase bridge without totals', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A', 'B', 'C'],
      series: [series({ values: [10, 20, 30] })],
      subtotalIndices: [],
      valAxisMajorGridlines: true,
      valAxisGridlineColor: 'D9D9D9',
      valAxisLineColor: '000000',
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(['A', 'B', 'C']);
    expect(rec.segs.some(segment =>
      segment.ss.toLowerCase() === '#d9d9d9'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100
    )).toBe(true);
    const gridline = rec.segs.find(segment =>
      Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100
      && segment.ss.toLowerCase() === '#d9d9d9'
      && segment.y0 > RECT.y
    );
    // Office keeps a visible axis rule off the chart-object edge even when its
    // implicit numeric labels are suppressed for this Waterfall layout.
    expect(gridline?.x0).toBeGreaterThanOrEqual(RECT.w * 0.03 - 0.01);
  });

  it('draws authored waterfall minor ticks', () => {
    const count = (minorTick: ChartModel['valAxisMinorTickMark']): number => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'waterfall', categories: ['Start', 'End'],
        series: [series({ values: [3, 23] })], subtotalIndices: [1],
        valMin: 3, valMax: 23, valAxisMajorUnit: 10, valAxisMinorUnit: 4,
        valAxisMajorGridlines: false, valAxisMinorGridlines: false,
        valAxisMajorTickMark: 'none', valAxisMinorTickMark: minorTick,
      }), RECT, 1);
      return rec.segs.filter(segment =>
        Math.abs(segment.y1 - segment.y0) < 0.01
        && Math.abs(segment.x1 - segment.x0) > 0
        && Math.abs(segment.x1 - segment.x0) <= 12
      ).length;
    };
    expect(count('cross') - count('none')).toBe(4);
  });
});

describe('ChartEx flat layouts dispatch to semantic renderers', () => {
  it('fails closed for an unknown future ChartEx layout without guessing from its data', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'futureLayout',
      categories: ['A', 'B'],
      series: [series({ values: [2, -1] })],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(['Unsupported chart']);
    expect(rec.rects).toHaveLength(0);
    expect(rec.strokeRects).toHaveLength(0);
    expect(rec.gradients).toHaveLength(0);
  });

  it('keeps unsupported-layout placeholder work constant for an unbounded public identifier', () => {
    const rec = recordingCtx();
    const chartType = 'x'.repeat(1_000_000);
    renderChart(rec.ctx, baseModel({ chartType, series: [series({ values: [1] })] }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(['Unsupported chart']);
  });

  it('measures the same semantic ChartEx column legend that it paints', () => {
    const renderPlot = (extraSeries: ChartSeries[]): RectCall => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredColumn',
        categories: ['A'],
        series: [
          series({ name: 'Bar', values: [1] }),
          ...extraSeries,
        ],
        showLegend: true,
        legendPos: 't',
        plotAreaBg: 'ABCDEF',
        chartexDataPointStyle: { fillColors: ['4472C4'] },
      }), { x: 0, y: 0, w: 240, h: 200 }, 1);
      return rec.rects.find(rect => rect.fs.toUpperCase() === '#ABCDEF') as RectCall;
    };

    const semanticOnly = renderPlot([]);
    const withNonLegendLine = renderPlot([
      series({
        name: 'A line series name that must not participate in the ChartEx column legend reserve',
        values: [1],
        seriesType: 'line',
      }),
    ]);
    expect(withNonLegendLine).toEqual(semanticOnly);
  });

  it('resolves clustered-column automatic style colors by series, not category', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn', categories: ['A', 'B', 'C'],
      showLegend: true,
      legendPos: 'r',
      series: [series({
        name: 'Histogram', values: [3, 2, 1], chartexFormatIdx: 2,
      })],
      chartexDataPointStyle: { fillColors: ['1F6A85', 'ED7D31', '70AD47'] },
    }), RECT, 1);

    expect(rec.rects.map(rect => rect.fs.toUpperCase())).toEqual([
      '#70AD47', '#70AD47', '#70AD47', '#70AD47',
    ]);
  });

  it('keeps non-contiguous ChartEx series format colors aligned with legend swatches', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn', categories: ['A', 'B'],
      showLegend: true,
      legendPos: 'r',
      series: [
        series({ name: 'First', values: [3, 2], chartexFormatIdx: 0 }),
        series({ name: 'Third format', values: [1, 4], chartexFormatIdx: 2 }),
      ],
      chartexDataPointStyle: { fillColors: ['1F6A85', 'ED7D31', '70AD47'] },
    }), RECT, 1);

    const colors = rec.rects.map(rect => rect.fs.toUpperCase());
    expect(colors.filter(color => color === '#1F6A85')).toHaveLength(3);
    expect(colors.filter(color => color === '#70AD47')).toHaveLength(3);
    expect(colors).not.toContain('#ED7D31');
  });

  it('keeps an explicit ChartEx data-point noFill transparent in the legend', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn', categories: ['A', 'B'],
      showLegend: true,
      legendPos: 'r',
      series: [series({ name: 'Transparent', values: [3, 2] })],
      chartexDataPointStyle: {
        fillHidden: true,
        fillNoStyle: false,
        lineColors: ['1F6A85'],
      },
    }), RECT, 1);

    expect(rec.rects).toHaveLength(0);
    expect(rec.strokeRects).toHaveLength(3);
  });

  it('applies clustered-column point paint and indexed series labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn', categories: ['A', 'B'],
      series: [
        series({ name: 'First', values: [1, 2] }),
        series({
          name: 'Second', values: [3, 4],
          dataPointOverrides: [
            { idx: 0, color: '112233', fillHidden: false, lineColor: '778899', lineWidthEmu: 25400, lineDash: 'dash', lineHidden: false },
            { idx: 1, fillHidden: true, lineHidden: true },
          ],
          dataLabelColors: ['445566', null],
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
            formatCode: '0.0', position: 'outEnd',
          },
          dataLabelOverrides: [
            { idx: 0, text: '', formatCode: '0.00', fontColor: '445566' },
            { idx: 1, text: '', deleted: true },
          ],
        }),
      ],
      chartexDataPointStyle: { fillColors: ['5B9BD5'], lineColors: ['FFFFFF'] },
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs.toUpperCase() === '#112233')).toHaveLength(1);
    expect(rec.rects).toHaveLength(3);
    expect(rec.strokeRects.some(rect => rect.ss.toUpperCase() === '#778899')).toBe(true);
    expect(rec.texts).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: '3.00', fillStyle: '#445566' }),
    ]));
    expect(rec.texts.map(text => text.text)).not.toContain('4.0');
  });

  it('bounds clustered-column primitives across multiple visible series', () => {
    const rec = recordingCtx();
    const values = Array.from({ length: 6_000 }, () => 1);
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn', categories: [],
      series: [series({ values }), series({ values })],
    }), RECT, 1);
    expect(rec.rects).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it.each([
    'clusteredBar', 'clusteredBarH', 'stackedBar',
    'line', 'area', 'pie', 'doughnut', 'radar',
    'scatter', 'bubble', 'stock',
  ])('bounds oversized classic %s input before family-specific layout or paint', chartType => {
    const rec = recordingCtx();
    const categories = Array.from({ length: 10_001 }, (_, index) => String(index));
    const values = Array.from({ length: 10_001 }, (_, index) => index + 1);
    renderChart(rec.ctx, baseModel({
      chartType,
      categories,
      series: [series({ categories, values })],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
    expect(rec.rects).toHaveLength(0);
    expect(rec.arcs).toHaveLength(0);
  });

  it.each(['clusteredColumn', 'funnel', 'paretoLine'])(
    '%s does not fall back to the unsupported-chart label',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: chartType === 'paretoLine' ? [] : ['1', '2', '3'],
        series: [series({ name: 'Authored series', values: [3, 2, 1] })],
        showLegend: true,
        legendPos: 't',
      }), RECT, 1);

      expect(rec.texts.map(text => text.text)).not.toContain(`Chart: ${chartType}`);
      if (chartType !== 'paretoLine') {
        expect(rec.texts.map(text => text.text)).toContain('Authored series');
      }
    },
  );

  it('bins raw histogram observations before routing them to bar geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'histogram',
      categories: [],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({ values: [0, 1, 2, 3, 4] })],
      chartexHistogramBinning: { binCount: 2, intervalClosed: 'l' },
    }), RECT, 1);

    expect(rec.rects).toHaveLength(2);
  });

  it('keeps ChartEx histogram value labels at the observed axis-relative offset', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'histogram',
      categories: [],
      series: [series({ values: [0, 1, 2, 3, 4] })],
      chartexHistogramBinning: { binCount: 2, intervalClosed: 'l' },
      valMin: 0,
      valMax: 4,
      valAxisMajorUnit: 2,
      valAxisFontSizeHpt: 1000,
      valAxisLineColor: '123456',
      valAxisLineHidden: false,
    }), RECT, 1);

    const axis = rec.segs.find(segment =>
      segment.ss === '#123456'
      && Math.abs(segment.x0 - segment.x1) < 0.001
      && Math.abs(segment.y1 - segment.y0) > 100
    );
    const zero = rec.texts.find(text => text.text === '0' && text.align === 'right');
    expect(axis).toBeDefined();
    expect(zero).toBeDefined();
    expect(axis!.x0 - zero!.x).toBeCloseTo(7, 5);
  });

  it('keeps dense one- and two-digit category labels on one line across fallback font metrics', () => {
    const categories = Array.from({ length: 29 }, (_, index) => String(index + 1));
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn',
      categories,
      catAxisFontSizeHpt: 1000,
      series: [series({ name: 'cyl', values: categories.map(() => 6) })],
    }), { x: 0, y: 0, w: 494, h: 288 }, 4 / 3);

    const categoryTexts = rec.texts
      .filter(text => text.baseline === 'top')
      .map(text => text.text);
    expect(categoryTexts).toEqual(expect.arrayContaining(categories));
  });

  it('rejects histogram input beyond the ChartEx cache ceiling', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'histogram',
      categories: [],
      series: [series({ values: new Array<number | null>(1_048_577) })],
    }), RECT, 1);

    expect(rec.rects).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it.each(['waterfall', 'funnel'])(
    '%s renders every numeric data point when the optional category dimension is unavailable',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: [],
        series: [series({ values: [3, 2, 1] })],
      }), RECT, 1);

      expect(rec.rects).toHaveLength(3);
    },
  );

  it.each(['waterfall', 'funnel'])(
    '%s routes authored labels through the shared four-line bounded layout',
    chartType => {
      const rec = recordingCtx();
      const authored = `A  ${'word '.repeat(1200)}`;
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        catAxisHidden: true,
        valAxisHidden: true,
        series: [series({
          values: [10],
          dataLabelOverrides: [{ idx: 0, text: authored, position: 'ctr', fontColor: '123456' }],
        })],
      }), RECT, 1);

      const labels = rec.texts.filter(text => text.fillStyle === '#123456');
      expect(labels.length).toBeGreaterThan(0);
      expect(labels.length).toBeLessThanOrEqual(4);
      expect(labels.map(text => text.text).join('')).toContain('A  ');
      expect(labels.map(text => text.text)).not.toContain(authored);
    },
  );

  it('applies ChartEx funnel manual label layout in chart coordinates', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'funnel',
      categories: ['A'],
      catAxisHidden: true,
      series: [series({
        values: [10],
        dataLabelOverrides: [{
          idx: 0,
          text: 'manual funnel',
          fontColor: '123456',
          manualLayout: { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.2, w: 0.2, h: 0.1 },
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(text => text.text === 'manual funnel')).toMatchObject({ x: 384, y: 90 });
  });

  it('keeps a point-level waterfall fontBold=false authoritative', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A'],
      dataLabelFontBold: true,
      series: [series({
        values: [10],
        dataLabelOverrides: [{ idx: 0, text: 'not bold', fontBold: false }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(text => text.text === 'not bold')?.font).not.toMatch(/^bold /);
  });

  it.each(['waterfall', 'funnel'] as const)(
    '%s honors a series-local data-label font face',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        catAxisHidden: true,
        valAxisHidden: true,
        series: [series({
          values: [10],
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
            position: 'ctr', fontFace: 'ChartEx Label Face',
          },
        })],
      }), RECT, 1);

      expect(rec.texts.find(text => text.text === '10')?.font)
        .toContain('"ChartEx Label Face"');
    },
  );

  it.each(['waterfall', 'funnel'])(
    '%s preserves category-only point slots when the string dimension is longer',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ values: [3] })],
      }), RECT, 1);

      expect(rec.rects).toHaveLength(3);
      expect(rec.texts.map(text => text.text)).toEqual(
        expect.arrayContaining(['A', 'B', 'C']),
      );
    },
  );

  it.each(['waterfall', 'funnel', 'clusteredColumn'])(
    '%s handles a large numeric dimension without variadic-argument overflow',
    chartType => {
      const values = Array.from({ length: 150_000 }, (_, index) => (index % 5) + 1);
      const rec = recordingCtx();
      expect(() => renderChart(rec.ctx, baseModel({
        chartType,
        categories: [],
        catAxisHidden: true,
        valAxisHidden: true,
        series: [series({
          values,
          seriesDataLabels: {
            showVal: false,
            showCatName: false,
            showSerName: false,
            showPercent: false,
          },
        })],
      }), { x: 0, y: 0, w: 320, h: 180 }, 1)).not.toThrow();
      expect(rec.rects).toHaveLength(0);
      expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
    },
  );

  it('keeps the Pareto category-axis rule while suppressing ordinal labels', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'paretoLine',
      categories: [],
      series: [series({ values: [3, 2, 1] })],
      catAxisLineColor: 'C00000',
      catAxisMajorTickMark: 'none',
    }), RECT, 1);

    expect(rec.segs.some(segment =>
      segment.ss.toLowerCase() === '#c00000' && segment.y0 === segment.y1
    )).toBe(true);
  });

  it('uses the ordinary linear value axis for a standalone Pareto line', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'paretoLine',
      categories: [],
      series: [series({ values: [5, 3, 2] })],
    }), { x: 0, y: 0, w: 371, h: 198 }, 1);

    const texts = rec.texts.map(text => text.text);
    expect(texts).toEqual(expect.arrayContaining(['0', '0.2', '0.4', '0.6', '0.8', '1', '1.2']));
    expect(texts.some(text => text.includes('%'))).toBe(false);
  });

  it('renders owner-backed Pareto bars in sorted source identity order with a 0-100% line axis', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pareto',
      categories: ['Five', 'Twenty', 'Ten'],
      series: [
        series({
          name: 'Frequency',
          values: [5, 20, 10],
          dataPointOverrides: [
            { idx: 0, color: 'AA0000' },
            { idx: 1, color: '00AA00' },
            { idx: 2, color: '0000AA' },
          ],
        }),
        series({ name: 'Cumulative %', values: [], color: '333333' }),
      ],
    }), RECT, 1);

    expect(rec.rects.map(rect => rect.fs.toUpperCase())).toEqual([
      '#00AA00', '#0000AA', '#AA0000',
    ]);
    const texts = rec.texts.map(text => text.text);
    expect(texts).toEqual(expect.arrayContaining(['Twenty', 'Ten', 'Five', '0%', '100%']));
  });

  it.each(['pareto', 'paretoLine'])('%s rejects oversized input before sorting or paint', chartType => {
    const rec = recordingCtx();
    const values = Array.from({ length: 10_001 }, (_, index) => index);
    renderChart(rec.ctx, baseModel({
      chartType,
      categories: [],
      series: [series({ values })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
    expect(rec.rects).toHaveLength(0);
  });

  it('standalone Pareto line resolves direct and linked ChartEx strokes', () => {
    const model = (
      direct: ChartSeries['chartexStyle'],
      linked: ChartModel['chartexDataPointLineStyle'],
    ) => baseModel({
      chartType: 'paretoLine',
      categories: ['A', 'B', 'C'],
      chartexDataPointLineStyle: linked,
      series: [series({ values: [3, 2, 1], chartexStyle: direct })],
    });
    const cumulativeStroke = (rec: ReturnType<typeof strokedPolylineCtx>) =>
      rec.strokes.find(stroke => stroke.points.some((point, index) =>
        index > 0
        && point.x !== stroke.points[index - 1].x
        && point.y !== stroke.points[index - 1].y
      ));

    const direct = strokedPolylineCtx();
    renderChart(direct.ctx, model(
      {
        lineColors: ['123456'], lineWidthEmu: 25400, lineDash: 'dash',
        lineCap: 'rnd', lineJoin: 'bevel',
      },
      { lineColors: ['AA0000'], lineWidthEmu: 12700 },
    ), RECT, 1);
    expect(cumulativeStroke(direct)).toMatchObject({
      ss: '#123456', lw: 2, cap: 'round', join: 'bevel',
    });
    expect(cumulativeStroke(direct)?.dash.length).toBeGreaterThan(0);

    const linkedNoFill = strokedPolylineCtx();
    renderChart(linkedNoFill.ctx, model(null, { lineHidden: true }), RECT, 1);
    expect(cumulativeStroke(linkedNoFill)).toBeUndefined();

    const directNoFill = strokedPolylineCtx();
    renderChart(
      directNoFill.ctx,
      model({ lineHidden: true }, { lineHidden: true, lineNoStyle: true }),
      RECT,
      1,
    );
    expect(cumulativeStroke(directNoFill)).toBeUndefined();

    const linkedNoStyle = strokedPolylineCtx();
    renderChart(linkedNoStyle.ctx, model(null, { lineHidden: true, lineNoStyle: true }), RECT, 1);
    expect(cumulativeStroke(linkedNoStyle)).toBeDefined();
  });

  it('honors direct Pareto cumulative-line width/color and noFill', () => {
    const model = (chartexStyle: ChartSeries['chartexStyle']) => baseModel({
      chartType: 'pareto',
      categories: ['A', 'B', 'C'],
      series: [
        series({ values: [3, 2, 1] }),
        series({
          values: [],
          seriesType: 'line',
          useSecondaryAxis: true,
          showMarker: false,
          chartexStyle,
        }),
      ],
    });

    const styled = strokedPolylineCtx();
    renderChart(styled.ctx, model({ lineColors: ['123456'], lineWidthEmu: 25400 }), RECT, 1);
    expect(styled.strokes.some(stroke =>
      stroke.ss.toLowerCase() === '#123456'
      && stroke.lw === 2
      && stroke.points.some((point, index) =>
        index > 0
        && point.x !== stroke.points[index - 1].x
        && point.y !== stroke.points[index - 1].y
      )
    )).toBe(true);

    const hidden = strokedPolylineCtx();
    renderChart(hidden.ctx, model({ lineHidden: true }), RECT, 1);
    expect(hidden.strokes.some(stroke =>
      stroke.points.some((point, index) =>
        index > 0
        && point.x !== stroke.points[index - 1].x
        && point.y !== stroke.points[index - 1].y
      )
    )).toBe(false);
  });

  it('distinguishes linked NoStyle from linked and direct Pareto line noFill', () => {
    const model = (
      linked: ChartModel['chartexDataPointLineStyle'],
      direct?: ChartSeries['chartexStyle'],
    ) => baseModel({
      chartType: 'pareto',
      categories: ['A', 'B', 'C'],
      chartexDataPointLineStyle: linked,
      series: [
        series({ values: [3, 2, 1] }),
        series({
          values: [],
          seriesType: 'line',
          useSecondaryAxis: true,
          showMarker: false,
          chartexStyle: direct,
        }),
      ],
    });
    const hasCumulativeLine = (rec: ReturnType<typeof strokedPolylineCtx>): boolean =>
      rec.strokes.some(stroke => stroke.points.some((point, index) =>
        index > 0
        && point.x !== stroke.points[index - 1].x
        && point.y !== stroke.points[index - 1].y
      ));

    const noStyle = strokedPolylineCtx();
    renderChart(noStyle.ctx, model({ lineHidden: true, lineNoStyle: true }), RECT, 1);
    expect(hasCumulativeLine(noStyle)).toBe(true);

    const linkedNoFill = strokedPolylineCtx();
    renderChart(linkedNoFill.ctx, model({ lineHidden: true }), RECT, 1);
    expect(hasCumulativeLine(linkedNoFill)).toBe(false);

    const directNoFill = strokedPolylineCtx();
    renderChart(
      directNoFill.ctx,
      model({ lineHidden: true, lineNoStyle: true }, { lineHidden: true }),
      RECT,
      1,
    );
    expect(hasCumulativeLine(directNoFill)).toBe(false);
  });

  it('uses the original combo-series index for linked line style fallback', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      chartexDataPointLineStyle: { lineColors: ['111111', '222222'] },
      series: [
        series({ values: [3, 2, 1] }),
        series({ values: [1, 2, 3], seriesType: 'line', showMarker: false }),
      ],
    }), RECT, 1);

    expect(rec.strokes.some(stroke =>
      stroke.ss.toLowerCase() === '#222222'
      && stroke.points.some((point, index) =>
        index > 0
        && point.x !== stroke.points[index - 1].x
        && point.y !== stroke.points[index - 1].y
      )
    )).toBe(true);
  });

  it('waterfall uses locale-neutral English semantic legend labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['1', '2', '3'],
      series: [series({ values: [3, -1, 2] })],
      subtotalIndices: [2],
      showLegend: true,
      legendPos: 't',
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['Increase', 'Decrease', 'Total']),
    );
  });

  it('waterfall preserves missing slots without painting or labeling non-finite values', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Invalid', 'Missing', 'Drop', 'End'],
      series: [series({ values: [2, Number.NaN, null, -1, 1] })],
      subtotalIndices: [4],
      showDataLabels: true,
      catAxisHidden: true,
      valAxisHidden: true,
    }), RECT, 1);

    // Three finite bars plus the historical zero-height placeholder for the
    // missing numeric slot. The present NaN point is the only suppressed bar.
    expect(rec.rects).toHaveLength(4);
    expect(rec.texts.map(text => text.text)).toEqual(['2', '-1', '1']);
  });

  it('waterfall keeps finite geometry when cumulative finite values overflow', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A', 'B', 'C'],
      series: [series({ values: [Number.MAX_VALUE, Number.MAX_VALUE, -1] })],
      catAxisHidden: true,
      valAxisHidden: true,
    }), RECT, 1);
    expect(rec.rects).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(chart values out of range)');
  });

  it('waterfall preserves an authored category-only label on a missing numeric slot', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Missing', 'End'],
      series: [series({
        values: [2, null, 2],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
        },
        dataLabelOverrides: [{
          idx: 1,
          text: '',
          showVal: false,
          showCatName: true,
        }],
      })],
      catAxisHidden: true,
      valAxisHidden: true,
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toContain('Missing');
    expect(rec.texts.map(text => text.text)).not.toContain('0');
  });

  it('measures the same semantic waterfall entries that it paints', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [3, -1, 2] })],
      subtotalIndices: [2],
      showLegend: true,
      legendPos: 't',
    }), { x: 0, y: 0, w: 150, h: 200 }, 1);

    const semanticLabels = rec.texts
      .map(text => text.text)
      .filter(text => ['Increase', 'Decrease', 'Total'].includes(text));
    expect(semanticLabels).toEqual(['Increase', 'Decrease', 'Total']);
  });

  it('keeps Waterfall point fill/outline formatting above series noFill and linked colors', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['Start', 'Increase', 'Decrease A', 'Decrease B', 'Decrease C', 'End'],
      subtotalIndices: [5],
      catAxisHidden: true,
      valAxisHidden: true,
      chartexDataPointStyle: {
        fillColors: ['E46970', '8977D7', 'A5A5A5'],
        fillPaintAuthored: true,
      },
      series: [series({
        values: [245, 235, -52, -40, -108, 280],
        chartexStyle: {
          fillHidden: true,
          fillPaintAuthored: true,
          lineColors: ['E46970'],
          linePaintAuthored: true,
        },
        dataPointOverrides: [
          { idx: 0, color: '196ECA', lineHidden: true },
          { idx: 1, lineColor: '196ECA' },
          { idx: 5, color: '196ECA', lineHidden: true },
        ],
      })],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#196ECA')).toHaveLength(2);
    expect(rec.rects.filter(rect => ['#E46970', '#8977D7', '#A5A5A5'].includes(rect.fs)))
      .toHaveLength(0);
    expect(rec.strokeRects.filter(rect => rect.ss === '#196ECA')).toHaveLength(1);
    expect(rec.strokeRects.filter(rect => rect.ss === '#E46970')).toHaveLength(3);
  });

  const chartExLegendModel = (
    chartType: string,
    localStyle: ChartSeries['chartexStyle'],
    linkedStyle: ChartModel['chartexDataPointStyle'] = {
      lineColors: ['AA0000'], lineWidthEmu: 12700,
    },
  ): ChartModel => {
    const owner = series({
      name: 'Authored', values: [3, 2, 1], chartexStyle: localStyle,
    });
    const common = {
      showLegend: true,
      legendPos: 'r' as const,
      chartexDataPointStyle: linkedStyle,
    };
    switch (chartType) {
      case 'histogram':
        return baseModel({
          ...common, chartType, series: [owner], chartexHistogramBinning: { binCount: 2 },
        });
      case 'waterfall':
        return baseModel({
          ...common, chartType, categories: ['A', 'B', 'C'], series: [owner], subtotalIndices: [2],
        });
      case 'funnel':
        return baseModel({
          ...common, chartType, categories: ['A', 'B', 'C'], series: [owner],
        });
      case 'boxWhisker':
        return baseModel({
          ...common,
          chartType,
          series: [series({ name: 'Authored', values: [] })],
          chartexBox: {
            categories: ['A'],
            series: [{
              name: 'Authored', color: '4472C4', chartexStyle: localStyle,
              valuesByCategory: [[1, 2, 3]], meanMarker: false, meanLine: false,
              showOutliers: false, showNonoutliers: false, quartileMethod: 'inclusive',
            }],
          },
        });
      case 'sunburst':
        return baseModel({
          ...common,
          chartType,
          series: [owner],
          chartexSunburst: { rows: [{ path: ['Branch', 'Leaf'], size: 3 }] },
        });
      case 'treemap':
        return baseModel({
          ...common,
          chartType,
          series: [owner],
          chartexTreemap: {
            parentLabelLayout: 'none', rows: [{ path: ['Branch', 'Leaf'], size: 3 }],
          },
        });
      default:
        return baseModel({
          ...common, chartType: 'clusteredColumn', categories: ['A'], series: [owner],
        });
    }
  };

  it.each([
    'clusteredColumn', 'histogram', 'waterfall', 'funnel',
    'boxWhisker',
  ])('%s legend uses the same direct ChartEx outline as its plotted mark', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, chartExLegendModel(chartType, {
      lineColors: ['123456'], lineWidthEmu: 25400, lineDash: 'dash',
      lineCap: 'rnd', lineJoin: 'bevel',
    }), RECT, 1);

    const keys = rec.strokeRects.filter(rect => rect.w <= 7 && rect.h <= 7);
    expect(keys.length).toBeGreaterThan(0);
    expect(keys.every(rect =>
      rect.ss.toLowerCase() === '#123456'
      && rect.lw === 2
      && rect.dash.length > 0
      && rect.cap === 'round'
      && rect.join === 'bevel'
    )).toBe(true);
  });

  it.each([
    'sunburst', 'treemap',
  ])('%s legend does not inherit the hierarchy separator outline', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, chartExLegendModel(chartType, {
      lineColors: ['FFFFFF'], lineWidthEmu: 12700,
    }), RECT, 1);

    expect(rec.strokeRects.filter(rect => rect.w <= 7 && rect.h <= 7)).toHaveLength(0);
  });

  it.each([
    'clusteredColumn', 'histogram', 'waterfall', 'funnel',
    'boxWhisker', 'sunburst', 'treemap',
  ])('%s legend keeps a direct ChartEx outline noFill authoritative', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, chartExLegendModel(chartType, { lineHidden: true }), RECT, 1);
    expect(rec.strokeRects.filter(rect => rect.w <= 7 && rect.h <= 7)).toHaveLength(0);
  });

  it.each([
    'clusteredColumn', 'histogram', 'waterfall', 'funnel',
    'boxWhisker', 'sunburst', 'treemap',
  ])('%s paints the shared plot-area frame behind ChartEx geometry', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...chartExLegendModel(chartType, null),
      plotAreaBg: '123456',
      plotAreaFillPaintAuthored: true,
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
        x: 0.2, y: 0.25, w: 0.4, h: 0.35,
      },
    }, RECT, 1);
    const plotFrame = rec.rects.find(rect => rect.fs === '#123456');
    expect(plotFrame).toBeDefined();
    expect(plotFrame?.x).toBeCloseTo(128);
    expect(plotFrame?.y).toBeCloseTo(90);
    expect(plotFrame?.w).toBeCloseTo(256);
    expect(plotFrame?.h).toBeCloseTo(126);
  });

  it.each([
    ['clusteredColumn', false],
    ['histogram', false],
    ['waterfall', false],
    ['funnel', false],
    ['boxWhisker', true],
    ['sunburst', false],
    ['treemap', false],
  ] as const)('%s legend preserves its semantic rule for linked NoStyle: %s', (chartType, expected) => {
    const rec = recordingCtx();
    renderChart(
      rec.ctx,
      chartExLegendModel(chartType, null, { lineHidden: true, lineNoStyle: true }),
      RECT,
      1,
    );
    expect(rec.strokeRects.some(rect => rect.w <= 7 && rect.h <= 7)).toBe(expected);
  });

  it.each([false, true])('keeps a point outline visible on a noFill %s-D pie legend key', useThreeD => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: ['Outlined'], showLegend: true,
      legendPos: 'r',
      threeD: useThreeD ? { rotationX: 15, rotationY: 20 } : undefined,
      series: [series({
        categories: ['Outlined'], values: [1], dataPointColors: ['00000000'],
        dataPointOverrides: [{
          idx: 0, fillHidden: true, lineColor: '00FF00', lineWidthEmu: 25_400,
          lineDash: 'dash',
        }],
      })],
    }), RECT, 1);
    expect(rec.strokeRects.some(rect =>
      rect.ss === '#00FF00' && rect.lw === 2 && rect.dash.length > 0
    )).toBe(true);
  });

  it('keeps a 2-D pie series noFill transparent and does not invent slice borders', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: ['A', 'B'],
      series: [series({
        color: '00000000', values: [1, 1], lineHidden: true,
        dataPointColors: [null, null],
      })],
    }), RECT, 1);
    expect(rec.paintEvents.some(event =>
      event.kind === 'fill' && event.fillStyle !== '#00000000'
    )).toBe(false);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle.toLowerCase() === '#fff'
    )).toBe(false);
  });

  it('uses the single series fill when 2-D pie varyColors is disabled', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', categories: ['A', 'B'], varyColors: false,
      series: [series({ color: '123456', values: [1, 1], dataPointColors: [null, null] })],
    }), RECT, 1);
    const fills = rec.paintEvents.filter(event => event.kind === 'fill');
    expect(fills.filter(event => event.fillStyle === '#123456')).toHaveLength(2);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle.toLowerCase() === '#fff'
    )).toBe(false);
  });
});

describe('scatter series data labels honor c:date1904 (§21.2.2.38)', () => {
  // The scatter path was the one call site (of 18) that did not thread
  // chart.date1904 into its data-label value formatter, so a date-format-code
  // label rendered against the 1900 epoch even in a 1904 chart (1462 days off).
  const SERIAL = 45292; // 1900-system 2024-01-01
  function scatterWithDateLabel(date1904: boolean): TextCall[] {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      date1904,
      series: [series({
        name: 'S',
        // No categories → useIndexX; the y-value carries the serial date.
        values: [SERIAL],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          formatCode: 'd-mmm-yy',
        },
      })],
    }), RECT, 1);
    return rec.texts;
  }

  it('formats the data label against the chart date system (1900 vs 1904 differ)', () => {
    const expected1900 = formatChartValWithCode(SERIAL, 'd-mmm-yy', false);
    const expected1904 = formatChartValWithCode(SERIAL, 'd-mmm-yy', true);
    // The two epochs are 1462 days apart, so the expected strings must differ —
    // otherwise the test could not tell whether date1904 was threaded.
    expect(expected1900).not.toBe(expected1904);

    expect(scatterWithDateLabel(false).some(t => t.text === expected1900)).toBe(true);
    expect(scatterWithDateLabel(true).some(t => t.text === expected1904)).toBe(true);
    // Guard against a regression that ignores the flag: the 1904 chart must NOT
    // emit the 1900-epoch label.
    expect(scatterWithDateLabel(true).some(t => t.text === expected1900)).toBe(false);
  });
});

// ─── CH7 — secondary value axis for line / area (§21.2.2.*) ──────────────────
//
// A combo can bind a series to a SECONDARY value axis (a second `<c:valAx>`
// with axPos="r" / `<c:crosses val="max">`). Bar already supports this; CH7
// extends it to the line and area families. The secondary series is plotted
// against the axis's OWN independent scale, and the axis is drawn on the right
// edge. Scatter is intentionally NOT wired (Excel/PowerPoint do not define a
// Y secondary axis for XY scatter).

/** Recording context that captures path vertices (moveTo/lineTo/arc) grouped
 *  into SEGMENTS delimited by `beginPath`, plus fillText. Line/area build each
 *  series as its own `beginPath`…path…`fill`/`stroke` sequence, so a segment
 *  isolates one series' plotted vertices — independent of when the renderer
 *  sets strokeStyle/fillStyle relative to the path ops (area sets them AFTER
 *  building the path, so strokeStyle-based grouping would misattribute). A test
 *  picks the segment for a series by its known draw order. `fillRect` is dropped
 *  (line/area draw no bars). */
function pathRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  segments: Array<Array<{ x: number; y: number }>>;
  texts: TextCall[];
} {
  const segments: Array<Array<{ x: number; y: number }>> = [];
  let current: Array<{ x: number; y: number }> | null = null;
  const texts: TextCall[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const push = (x: number, y: number): void => {
    if (!current) { current = []; segments.push(current); }
    current.push({ x, y });
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { current = null; };
        case 'moveTo':
        case 'lineTo':
        case 'arc':
          return (x: number, y: number) => push(x, y);
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segments, texts };
}

const SECONDARY_AXIS = {
  min: null,
  max: null,
  title: 'Rate',
  hidden: false,
  majorTickMark: 'out',
  lineHidden: false,
};

describe('axis noFill tick visibility', () => {
  it('suppresses value-axis tick lines with a hidden axis rule while retaining labels', () => {
    const render = (lineHidden: boolean) => {
      const rec = pathRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['2008', '2009', '2010'],
        series: [series({ values: [0.18, 0.19, 0.17] })],
        valAxisLineHidden: lineHidden,
        valAxisMajorTickMark: 'out',
        catAxisMajorTickMark: 'none',
        valAxisFormatCode: '0%',
      }), RECT, 1);
      return rec;
    };
    const visible = render(false);
    const hidden = render(true);
    const valueTicks = (segments: Array<Array<{ x: number; y: number }>>) => segments.filter((segment) =>
      segment.length === 2
      && segment[0].y === segment[1].y
      && Math.abs(segment[1].x - segment[0].x) <= 8,
    );

    expect(valueTicks(visible.segments).length).toBeGreaterThan(0);
    expect(valueTicks(hidden.segments)).toHaveLength(0);
    expect(hidden.texts.some((text) => text.text.endsWith('%'))).toBe(true);
  });
});

describe('CH7 — line/area series honor a secondary value axis (§21.2.2.*)', () => {
  // The primary series ASCENDS [10,20,30]; the secondary series DESCENDS
  // [3,2,1]. Opposite slopes make the secondary series identifiable by geometry
  // alone (no color/draw-order coupling): its plotted profile falls left→right,
  // the primary's rises. Crucially the secondary series peaks at the FIRST
  // category (value 3). Mapped to its OWN axis (0..~3.5) that peak rides near
  // the plot top; mapped to the PRIMARY axis (0..~35) value 3 barely leaves the
  // bottom. The primary series peaks at the LAST category, so the LEFT third of
  // the plot contains a high point ONLY when the secondary axis is wired.
  const primaryVals = [10, 20, 30];
  const secondaryVals = [3, 2, 1];

  function comboModel(chartType: 'line' | 'area', withSecondaryAxis: boolean): ChartModel {
    return baseModel({
      chartType,
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Big', values: primaryVals }),
        series({ name: 'Small', values: secondaryVals, useSecondaryAxis: true }),
      ],
      secondaryValAxis: withSecondaryAxis ? { ...SECONDARY_AXIS } : null,
    });
  }

  /** A "data" segment is a polyline/fill that slopes — its vertices vary in BOTH
   *  x and y. Gridlines (constant y) and axis rules (constant x) are flat in one
   *  axis, so this filter isolates the plotted series geometry from the chrome. */
  function isDataSegment(seg: Array<{ x: number; y: number }>): boolean {
    if (seg.length < 3) return false;
    const xs = new Set(seg.map(p => Math.round(p.x)));
    const ys = new Set(seg.map(p => Math.round(p.y)));
    return xs.size > 1 && ys.size > 1;
  }

  /** Highest (min-Y) DATA vertex in the LEFT third of the plot. The primary
   *  series' high point is on the RIGHT, so a high point here can only be the
   *  DESCENDING secondary series' value-3 peak — present only when that series
   *  rides its own (short) axis. Chrome (gridlines / axis rules) is excluded, so
   *  the measure reflects series geometry alone; independent of color/draw order. */
  function leftPeakY(segments: Array<Array<{ x: number; y: number }>>): number {
    const leftThird = RECT.x + RECT.w / 3;
    const ys = segments
      .filter(isDataSegment)
      .flat()
      .filter(p => p.x < leftThird)
      .map(p => p.y);
    expect(ys.length).toBeGreaterThan(0);
    return Math.min(...ys);
  }

  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: the secondary series maps to its OWN scale, not the primary`, () => {
      const wired = pathRecordingCtx();
      renderChart(wired.ctx, comboModel(chartType, true), RECT, 1);
      const unwired = pathRecordingCtx();
      renderChart(unwired.ctx, comboModel(chartType, false), RECT, 1);
      // Wired: the descending series' value-3 peak sits top-left (small Y).
      // Unwired: value 3 on the tall primary axis stays low, so the left third
      // has no high point — its min-Y is far larger. A ≥100px gap can't be noise.
      const wiredPeak = leftPeakY(wired.segments);
      const unwiredPeak = leftPeakY(unwired.segments);
      expect(wiredPeak).toBeLessThan(unwiredPeak - 100);
    });

    it(`${chartType}: draws right-edge secondary axis tick labels + title`, () => {
      const rec = pathRecordingCtx();
      renderChart(rec.ctx, comboModel(chartType, true), RECT, 1);
      // Primary value labels sit LEFT of the plot; secondary tick labels + title
      // sit to the RIGHT. A text mark past 75% of the width can only be secondary.
      const rightLabels = rec.texts.filter(t => t.x > RECT.x + RECT.w * 0.75);
      expect(rightLabels.length).toBeGreaterThan(0);
      expect(rec.texts.some(t => t.text === 'Rate')).toBe(true);
    });

    it(`${chartType}: NO secondary axis (secondaryValAxis null) → no right-edge labels/title`, () => {
      // Byte-stability guard: without a secondary axis the renderer must draw NO
      // right-edge axis marks — it degrades to the exact single-axis path.
      const rec = pathRecordingCtx();
      renderChart(rec.ctx, comboModel(chartType, false), RECT, 1);
      expect(rec.texts.some(t => t.text === 'Rate')).toBe(false);
    });
  }

  it('paints the right-axis title bottom-to-top with its authored font', () => {
    const rec = ringRecordingCtx();
    const model = comboModel('line', true);
    model.secondaryValAxis = {
      ...(model.secondaryValAxis as NonNullable<ChartModel['secondaryValAxis']>),
      title: 'Rate',
      titleFontFace: 'Aptos Narrow',
      titleFontSizeHpt: 900,
      titleFontBold: false,
    };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.rotates).toContain(-Math.PI / 2);
    const title = rec.fontTexts.find(text => text.text === 'Rate');
    expect(title?.font).toContain('9px');
    expect(title?.font).toContain('Aptos Narrow');
  });
});

// ─── CH9 — line/area marker detail, error bars, per-point labels, smooth,
//          dispBlanksAs (§21.2.2.32 / §21.2.2.20 / §21.2.2.45 / §21.2.2.194 /
//          §21.2.2.42) ─────────────────────────────────────────────────────
//
// scatter already consumes s.markerSymbol/size/fill/line, s.errBars,
// s.dataLabelOverrides + s.seriesDataLabels, and smooth splines. CH9 wires the
// same series-level fields into the line and area families, adds per-series
// smooth (`<c:ser><c:smooth>`), and honors the chartSpace `dispBlanksAs` value
// when deciding how null cells break/span/zero the plotted line.

interface ArcCall { x: number; y: number; r: number; fillStyle: string }
interface FillRectCall { x: number; y: number; w: number; h: number }

/** Recording context that captures the primitives markers / smooth / error
 *  bars emit: `arc` (circle/star markers + the default line dot), `fillRect`
 *  (square marker + dash), `bezierCurveTo` (smooth spline), and `fillText`
 *  (data labels). Also groups stroked/filled path vertices into SEGMENTS
 *  (delimited by `beginPath`) so a test can inspect the polyline a series
 *  drew — used to tell gap / zero / span apart for dispBlanksAs. */
function markerRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  arcs: ArcCall[];
  fillRects: FillRectCall[];
  fillCalls: number;
  beziers: number;
  texts: TextCall[];
  segments: Array<Array<{ x: number; y: number }>>;
  strokeStyles: string[];
} {
  const arcs: ArcCall[] = [];
  const fillRects: FillRectCall[] = [];
  const texts: TextCall[] = [];
  const segments: Array<Array<{ x: number; y: number }>> = [];
  const strokeStyles: string[] = [];
  let current: Array<{ x: number; y: number }> | null = null;
  let beziers = 0;
  let fillCalls = 0;
  const state: Record<string, unknown> = {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const push = (x: number, y: number): void => {
    if (!current) { current = []; segments.push(current); }
    current.push({ x, y });
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { current = null; };
        case 'moveTo':
        case 'lineTo':
          return (x: number, y: number) => push(x, y);
        case 'arc':
          return (x: number, y: number, rad: number) => {
            arcs.push({ x, y, r: rad, fillStyle: String(state.fillStyle) });
            push(x, y);
          };
        case 'fillRect':
          return (x: number, y: number, w: number, h: number) => fillRects.push({ x, y, w, h });
        case 'fill':
          return () => { fillCalls += 1; };
        case 'stroke':
          return () => { strokeStyles.push(String(state.strokeStyle)); };
        case 'bezierCurveTo':
          return () => { beziers += 1; };
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    arcs, fillRects, texts, segments, strokeStyles,
    get beziers() { return beziers; },
    get fillCalls() { return fillCalls; },
  } as never;
}

describe('CH9 — line/area consume marker detail (§21.2.2.32)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: markerSymbol="square" draws square markers (fillRect), not the default circle`, () => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [3, 5, 4], showMarker: true, markerSymbol: 'square' })],
      }), RECT, 1);
      // One square fillRect per data point. (Area also fills the region with a
      // path, not a fillRect, so every fillRect here is a marker.)
      expect(rec.fillRects.length).toBe(3);
      // Squares are square: w === h.
      for (const fr of rec.fillRects) expect(Math.round(fr.w)).toBe(Math.round(fr.h));
    });

    it(`${chartType}: markerSize scales the marker (bigger size → bigger square)`, () => {
      const small = markerRecordingCtx();
      renderChart(small.ctx, baseModel({
        chartType,
        categories: ['Alpha', 'Beta'],
        series: [series({ name: 'S', values: [3, 5], showMarker: true, markerSymbol: 'square', markerSize: 4 })],
      }), RECT, 1);
      const big = markerRecordingCtx();
      renderChart(big.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        series: [series({ name: 'S', values: [3, 5], showMarker: true, markerSymbol: 'square', markerSize: 20 })],
      }), RECT, 1);
      expect(big.fillRects[0].w).toBeGreaterThan(small.fillRects[0].w);
    });

    it(`${chartType}: a series WITHOUT markerSymbol keeps the default circle marker`, () => {
      // Byte-stability: the fixed-circle fast path must remain when no symbol
      // is specified — no fillRect (square), markers are drawn via arc.
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [3, 5, 4], showMarker: true })],
      }), RECT, 1);
      expect(rec.fillRects.length).toBe(0);
      // 3 marker dots (arcs). Line also strokes with arc-free paths, so all
      // arcs are markers here.
      const markerArcs = rec.arcs.filter(a => a.r < 10);
      expect(markerArcs.length).toBe(3);
    });
  }

  it('uses the authored marker-outline width independently from marker size', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3, 5],
        showMarker: true,
        markerSymbol: 'circle',
        markerSize: 7,
        markerFill: 'FFFFFF',
        markerLine: '1696D2',
        markerLineWidthEmu: 25400,
        lineColor: '1696D2',
        lineWidthEmu: 25400,
      })],
    }), RECT, 1);

    const blueStrokes = rec.strokes.filter(stroke =>
      stroke.strokeStyle.toLowerCase() === '#1696d2'
    );
    // The series path and both marker outlines share the authored 2pt stroke.
    expect(blueStrokes.length).toBe(3);
    expect(blueStrokes.every(stroke => stroke.lineWidth === 2)).toBe(true);
  });

  it('uses linked dataPointMarker paint behind direct classic-marker formatting', () => {
    const linked = ringRecordingCtx();
    renderChart(linked.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3, 5], showMarker: true, markerSymbol: 'circle', lineHidden: true,
      })],
      chartStyleRoles: {
        dataPointMarker: {
          fillColors: ['AABBCC'], lineColors: ['CCBBAA'], lineWidthEmu: 19_050,
        },
      },
    }), RECT, 1);
    expect(linked.fills.filter(fill => fill === '#AABBCC')).toHaveLength(2);
    expect(linked.strokes.filter(stroke =>
      stroke.strokeStyle === '#CCBBAA' && stroke.lineWidth === 1.5
    )).toHaveLength(2);

    const direct = ringRecordingCtx();
    renderChart(direct.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'circle', lineHidden: true,
        markerFill: '112233', markerLine: '332211',
      })],
      chartStyleRoles: {
        dataPointMarker: { fillHidden: true, lineHidden: true },
      },
    }), RECT, 1);
    expect(direct.fills).toContain('#112233');
    expect(direct.strokes.some(stroke => stroke.strokeStyle === '#332211')).toBe(true);
  });

  it('renders structured marker fills with direct point paint taking precedence', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3, 5], showMarker: true, markerSymbol: 'circle', lineHidden: true,
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'DDEEFF' },
          ],
        },
        dataPointOverrides: [{ idx: 1, markerFill: 'ABCDEF' }],
      })],
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(221,238,255,1)' },
    ]);
  });

  it('renders a picture marker from the host image lookup with authored source crop', () => {
    const rec = recordingCtx();
    const bitmap = { width: 100, height: 80 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 10,
        lineHidden: true,
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'xl/media/marker.png', mimeType: 'image/png',
          srcRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 },
        },
      })],
    }), RECT, 1, 0, testThreeD, undefined, fill => {
      expect(fill.imagePath).toBe('xl/media/marker.png');
      return bitmap;
    });
    expect(rec.drawImages).toHaveLength(1);
    expect(rec.drawImages[0].slice(0, 5)).toEqual([bitmap, 10, 16, 60, 56]);
    expect(rec.drawImages[0].slice(7)).toEqual([10, 10]);
  });

  it('keeps point, series, and linked picture-marker precedence', () => {
    const rec = recordingCtx();
    const linkedBitmap = { width: 40, height: 40 } as unknown as CanvasImageSource;
    const pointBitmap = { width: 60, height: 60 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      catAxisHidden: true, valAxisHidden: true,
      chartStyleRoles: {
        dataPointMarker: {
          fillPaintAuthored: true,
          fillPaints: [{
            fillType: 'image', stretch: true, imagePath: 'xl/media/linked.png', mimeType: 'image/png',
          }],
        },
      },
      series: [series({
        values: [3, 5], showMarker: true, markerSymbol: 'picture', lineHidden: true,
        dataPointOverrides: [{
          idx: 1, markerSymbol: 'picture', markerFillPaintAuthored: true,
          markerFillPaint: {
            fillType: 'image', stretch: true, imagePath: 'xl/media/point.png', mimeType: 'image/png',
          },
        }],
      })],
    }), RECT, 1, 0, testThreeD, undefined, fill =>
      fill.imagePath.endsWith('point.png') ? pointBitmap : linkedBitmap);

    expect(rec.drawImages.map(call => call[0])).toEqual([linkedBitmap, pointBitmap]);
  });

  it('reuses a picture marker for plot, legend, data-label keys, and data-table key', () => {
    const rec = recordingCtx();
    const bitmap = { width: 32, height: 32 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'], showLegend: true,
      series: [series({
        name: 'Pictures', values: [3, 5], showMarker: true, markerSymbol: 'picture',
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'xl/media/marker.png', mimeType: 'image/png',
        },
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false,
          showPercent: false, showLegendKey: true,
        },
      })],
      dataTable: {
        showHorizontalBorder: false, showVerticalBorder: false,
        showOutline: false, showKeys: true,
      },
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    // two plot points + one ordinary legend key + two data-label keys + one table key
    expect(rec.drawImages).toHaveLength(6);
    expect(rec.drawImages.every(call => call[0] === bitmap)).toBe(true);
  });

  it('renders picture markers through the optional 3-D marker path', () => {
    const rec = recordingCtx();
    const bitmap = { width: 32, height: 32 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        values: [3, 5], showMarker: true, markerSymbol: 'picture',
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'ppt/media/marker.png', mimeType: 'image/png',
        },
      })],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    expect(rec.drawImages).toHaveLength(2);
  });

  it('uses picture fill for optional 3-D dash markers in plot and legend', () => {
    const rec = recordingCtx();
    const bitmap = { width: 32, height: 32 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'], showLegend: true,
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        name: 'Dash pictures', values: [3, 5], showMarker: true, markerSymbol: 'dash', markerSize: 10,
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'ppt/media/dash-marker.png', mimeType: 'image/png',
        },
      })],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    // Two plot markers and the compound line/marker legend key all use the
    // authored fill. This also pins the shared collector/work fill predicate.
    expect(rec.drawImages).toHaveLength(3);
    expect(rec.clips.filter(clip => Math.abs(clip.h - 2) < 1e-6)).toHaveLength(2);
  });

  it.each([
    { threeD: false },
    { threeD: true },
  ])('uses the normative 1/2 by 1/5 dot geometry in $threeD marker paint', ({ threeD }) => {
    const rec = recordingCtx();
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      ...(threeD ? { threeD: { rotationX: 15, rotationY: 20, perspective: 30 } } : {}),
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'dot', markerSize: 10,
        markerFill: '00AA00', markerLine: 'FF0000', markerLineWidthEmu: 25_400,
      })],
    }), RECT, 1, 0, testThreeD);

    expect(rec.ellipses).toContainEqual(expect.objectContaining({ rx: 2.5, ry: 1 }));
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
  });

  it.each([
    { threeD: false },
    { threeD: true },
  ])('keeps a picture marker outline when authored fill is absent in $threeD paint', ({ threeD }) => {
    const rec = recordingCtx();
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      ...(threeD ? { threeD: { rotationX: 15, rotationY: 20, perspective: 30 } } : {}),
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 10,
        markerFill: '00000000', markerFillPaint: null, markerFillPaintAuthored: true,
        markerLine: 'FF0000', markerLineWidthEmu: 25_400,
      })],
    }), RECT, 1, 0, testThreeD);

    expect(rec.drawImages).toEqual([]);
    expect(rec.strokeRects).toContainEqual(expect.objectContaining({ ss: '#FF0000', w: 10, h: 10 }));
  });

  it('renders picture markers for an optional 3-D area series', () => {
    const rec = recordingCtx();
    const bitmap = { width: 32, height: 32 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'area', categories: ['A', 'B'],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        values: [3, 5], showMarker: true, markerSymbol: 'picture',
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'ppt/media/area-marker.png', mimeType: 'image/png',
        },
      })],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    expect(rec.drawImages).toHaveLength(2);
  });

  it.each([
    { chartType: 'line', seriesType: undefined },
    { chartType: 'area', seriesType: undefined },
    { chartType: 'clusteredBar', seriesType: 'line' },
    { chartType: 'clusteredBar', seriesType: 'area' },
  ] as const)('honors a point-only picture marker in $chartType/$seriesType', ({
    chartType, seriesType,
  }) => {
    const rec = recordingCtx();
    const bitmap = { width: 24, height: 24 } as unknown as CanvasImageSource;
    const overlay = series({
      name: 'Overlay', values: [3, 5], seriesType, showMarker: false,
      dataPointOverrides: [{
        idx: 0, markerSymbol: 'picture', markerSize: 10,
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'xl/media/point-only.png', mimeType: 'image/png',
        },
      }],
    });
    renderChartCore(rec.ctx, baseModel({
      chartType, categories: ['A', 'B'], catAxisHidden: true, valAxisHidden: true,
      series: seriesType
        ? [series({ name: 'Bars', values: [2, 4], seriesType: 'bar' }), overlay]
        : [overlay],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    expect(rec.drawImages).toHaveLength(1);
    expect(rec.drawImages[0][0]).toBe(bitmap);
  });

  it('tiles an authored picture marker within the marker clip', () => {
    const rec = recordingCtx();
    const bitmap = { width: 4, height: 4 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 10,
        markerFillPaint: {
          fillType: 'image', stretch: false, imagePath: 'xl/media/tile.png', mimeType: 'image/png',
          dpi: 96,
          tile: { tx: 0, ty: 0, sx: 0.5, sy: 0.5, flip: 'xy', algn: 'ctr' },
        },
      })],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    expect(rec.drawImages.length).toBeGreaterThan(4);
    expect(rec.clips.length).toBeGreaterThan(0);
  });

  it('uses authored blipFill dpi for tile physical size and counter-rotates only when requested', () => {
    const bitmap = { width: 300, height: 300 } as unknown as CanvasImageSource;
    const paint = (dpi: number, rotWithShape: boolean) => {
      const rec = recordingCtx();
      renderChartCore(rec.ctx, baseModel({
        chartType: 'line', categories: ['A'], catAxisHidden: true, valAxisHidden: true,
        series: [series({
          values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 20,
          markerFillPaint: {
            fillType: 'image', stretch: false, imagePath: 'xl/media/tile.png', mimeType: 'image/png',
            dpi, rotWithShape,
            tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
          },
        })],
      }), RECT, 1, 30, testThreeD, undefined, () => bitmap);
      return rec;
    };
    const dpi100 = paint(100, true);
    const dpi300 = paint(300, false);
    expect(Number(dpi100.drawImages[0][3]) / Number(dpi300.drawImages[0][3])).toBeCloseTo(3);
    expect(dpi100.rotations.some(value => Math.abs(value + Math.PI / 6) < 1e-9)).toBe(false);
    expect(dpi300.rotations.some(value => Math.abs(value + Math.PI / 6) < 1e-9)).toBe(true);
  });

  it('uses the same physical tile scale for plot and legend marker consumers', () => {
    const rec = recordingCtx();
    const bitmap = { width: 300, height: 300 } as unknown as CanvasImageSource;
    const picture = {
      fillType: 'image' as const, stretch: false,
      imagePath: 'xl/media/tile.png', mimeType: 'image/png', dpi: 300,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'], showLegend: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 150,
        markerFillPaint: picture,
      })],
    }), RECT, 4 / 3, 0, testThreeD, undefined, () => bitmap);
    expect(rec.drawImages.length).toBeGreaterThan(1);
    expect(rec.drawImages.every(call => Math.abs(Number(call[3]) - 96) < 1e-9)).toBe(true);
  });

  it('preserves negative srcRect outset space inside every image tile', () => {
    const rec = recordingCtx();
    const bitmap = { width: 100, height: 100 } as unknown as CanvasImageSource;
    const picture = {
      fillType: 'image' as const, stretch: false,
      imagePath: 'xl/media/outset-tile.png', mimeType: 'image/png', dpi: 96,
      srcRect: { l: -0.5, t: 0, r: 0, b: 0 },
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    renderChartCore(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'picture', markerSize: 20,
        markerFillPaint: picture,
      })],
    }), RECT, 1, 0, testThreeD, undefined, () => bitmap);

    const tiledDraws = rec.drawImages.filter(call => call.length === 9);
    expect(tiledDraws.length).toBeGreaterThan(0);
    expect(tiledDraws.every(call => Math.abs(Number(call[5]) * 2 - Number(call[7])) < 1e-9))
      .toBe(true);
    expect(tiledDraws.every(call => Number(call[1]) === 0 && Number(call[3]) === 100)).toBe(true);
  });

  it('applies the flip schema default but fails closed for unproven tile placement', () => {
    const bitmap = { width: 16, height: 16 } as unknown as CanvasImageSource;
    const lookup = () => bitmap;
    const base = {
      fillType: 'image' as const, stretch: false,
      imagePath: 'xl/media/tile.png',
      mimeType: 'image/png',
      dpi: 96,
    };
    const complete = { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' };
    for (const key of ['tx', 'ty', 'sx', 'sy', 'algn'] as const) {
      const tile: Partial<typeof complete> = { ...complete };
      delete tile[key];
      expect(chartImageFillPaintWork({ ...base, tile }, lookup, 20, 20)).toBe(0);
    }
    expect(chartImageFillPaintWork({
      ...base,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, algn: 'tl' },
    }, lookup, 20, 20)).toBeGreaterThan(0);
    expect(chartImageFillPaintWork({
      ...base,
      tile: { tx: 0, ty: 0, sx: -1, sy: 1, flip: 'none', algn: 'tl' },
    }, lookup, 20, 20)).toBe(0);
  });

  it('prefetches only effective picture-marker consumers', () => {
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/marker.png',
      mimeType: 'image/png',
    };
    const linked = {
      fillPaints: [picture], fillPaintAuthored: true,
    };
    const solidDirect = baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({ values: [1], showMarker: true, markerFill: 'FF0000' })],
      chartStyleRoles: { dataPointMarker: linked },
    });
    expect(collectChartMarkerImageFills(solidDirect)).toEqual([]);

    const noMode = { ...picture, stretch: undefined };
    const noModeModel = baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({ values: [1], showMarker: true, markerFillPaint: noMode })],
    });
    expect(collectChartMarkerImageFills(noModeModel)).toEqual([]);
    expect(chartImageFillPaintWork(noMode, () => ({ width: 1, height: 1 } as CanvasImageSource), 5, 5))
      .toBe(0);

    const fullyCropped = { ...picture, srcRect: { l: 0.6, t: 0, r: 0.6, b: 0 } };
    expect(collectChartMarkerImageFills(baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({ values: [1], showMarker: true, markerFillPaint: fullyCropped })],
    }))).toEqual([]);

    const legendOnly = baseModel({
      chartType: 'line', showLegend: true,
      series: [series({ values: [], markerSymbol: 'picture', markerFillPaint: picture })],
    });
    expect(collectChartMarkerImageFills(legendOnly)).toEqual([picture]);

    const tableFromSeriesCategories = baseModel({
      chartType: 'line', categories: [],
      dataTable: {
        showHorizontalBorder: false, showVerticalBorder: false,
        showOutline: false, showKeys: true,
      },
      series: [series({
        categories: ['A'], values: [1], markerSymbol: 'picture', markerFillPaint: picture,
      })],
    });
    expect(collectChartMarkerImageFills(tableFromSeriesCategories)).toEqual([picture]);

    const meanOnlyBox = baseModel({
      chartType: 'boxWhisker',
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', valuesByCategory: [[1, 2, 3]], meanMarker: true, meanLine: false,
          showOutliers: false, showNonoutliers: false, quartileMethod: 'exclusive',
          chartexStyle: { fillPaints: [picture], fillPaintAuthored: true },
        }],
      },
    });
    expect(collectChartMarkerImageFills(meanOnlyBox)).toEqual([]);
  });

  it('prefetches a series picture used only by data-label legend keys', () => {
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/label-key.png', mimeType: 'image/png',
    };
    const model = baseModel({
      chartType: 'line', categories: ['A', 'B'], showLegend: false,
      series: [series({
        values: [1, 2], showMarker: true, markerSymbol: 'picture', markerFillPaint: picture,
        dataPointOverrides: [
          { idx: 0, markerSymbol: 'none' }, { idx: 1, markerSymbol: 'none' },
        ],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false,
          showPercent: false, showLegendKey: true,
        },
      })],
    });
    expect(collectChartMarkerImageFills(model)).toEqual([picture]);
    const rec = recordingCtx();
    const bitmap = { width: 16, height: 16 } as unknown as CanvasImageSource;
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(rec.drawImages).toHaveLength(2);
  });

  it('rejects oversized public chart models before picture-prefetch traversal', () => {
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/oversized.png', mimeType: 'image/png',
    };
    const model = baseModel({
      chartType: 'line',
      series: [series({
        values: Array.from({ length: 10_001 }, () => 1),
        showMarker: true, markerSymbol: 'picture', markerFillPaint: picture,
      })],
    });
    expect(collectChartMarkerImageFills(model)).toEqual([]);
  });

  it('rejects global-category × series and label-override expansion before prefetch', () => {
    const categories = Array.from({ length: 5_001 }, (_, index) => String(index));
    const manySeries = baseModel({
      chartType: 'line', categories,
      series: Array.from({ length: 5_001 }, () => series({
        values: [], showMarker: true, markerSymbol: 'picture',
      })),
    });
    expect(collectChartMarkerImageFills(manySeries)).toEqual([]);

    const manyLabels = baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: Array.from({ length: 10_001 }, (_, idx) => ({
          idx, text: '', showLegendKey: true,
        })),
      })],
    });
    expect(collectChartMarkerImageFills(manyLabels)).toEqual([]);

    let legendIndexReads = 0;
    const boundedLegend = baseModel({
      chartType: 'line',
      series: Array.from({ length: 5_000 }, () => series({ values: [] })),
      legendEntries: Array.from({ length: 5_000 }, (_, index) => ({
        get idx() { legendIndexReads++; return index; },
        deleted: true,
      })),
    });
    expect(collectChartMarkerImageFills(boundedLegend)).toEqual([]);
    expect(legendIndexReads).toBe(5_000);

    const values = Array.from({ length: 100 }, () => 1);
    const trendlineWork = baseModel({
      chartType: 'line',
      series: [series({
        values,
        trendLines: Array.from({ length: 101 }, () => ({ trendlineType: 'linear' })),
      })],
    });
    expect(collectChartMarkerImageFills(trendlineWork)).toEqual([]);

    const errorBar = {
      dir: 'y', barType: 'both', plus: values, minus: values, noEndCap: false,
    };
    const errorBarWork = baseModel({
      chartType: 'line',
      series: [series({ values, errBars: Array.from({ length: 101 }, () => errorBar) })],
    });
    expect(collectChartMarkerImageFills(errorBarWork)).toEqual([]);
    const emptyErrorBars = baseModel({
      chartType: 'line',
      series: [series({
        values: [1],
        errBars: Array.from({ length: 10_001 }, () => ({
          dir: 'y', barType: 'both', plus: [], minus: [], noEndCap: false,
        })),
      })],
    });
    expect(collectChartMarkerImageFills(emptyErrorBars)).toEqual([]);
  });

  it('matches ChartEx box marker visibility and fill-layer precedence during prefetch', () => {
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/box.png', mimeType: 'image/png',
    };
    const box = (over: Record<string, unknown>) => baseModel({
      chartType: 'boxWhisker',
      chartexDataPointMarkerStyle: { fillPaints: [picture], fillPaintAuthored: true },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', valuesByCategory: [[1, 2, 3]], meanMarker: false, meanLine: false,
          showOutliers: false, showNonoutliers: true, quartileMethod: 'inclusive',
          ...over,
        }],
      },
    });
    expect(collectChartMarkerImageFills(box({
      chartexStyle: { fillHidden: true, fillNoStyle: true },
    }))).toEqual([picture]);
    expect(collectChartMarkerImageFills(box({
      chartexStyle: { fillHidden: true, fillNoStyle: false },
    }))).toEqual([]);
    expect(collectChartMarkerImageFills({
      ...box({}), chartStyleMarkerSymbol: 'none',
    })).toEqual([]);
    expect(collectChartMarkerImageFills(box({
      valuesByCategory: [[1, 2, 3]], showNonoutliers: false, showOutliers: true,
    }))).toEqual([]);
    expect(collectChartMarkerImageFills(box({
      valuesByCategory: [[1, 2, 3, 4, 100]], showNonoutliers: false, showOutliers: true,
    }))).toEqual([picture]);
  });

  it('atomically bounds unique decoded picture-marker sources', () => {
    const model = (count: number) => baseModel({
      chartType: 'line', categories: Array.from({ length: count }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: count }, () => 1), showMarker: false,
        dataPointOverrides: Array.from({ length: count }, (_, idx) => ({
          idx, markerSymbol: 'picture' as const,
          markerFillPaint: {
            fillType: 'image' as const, stretch: true,
            imagePath: `xl/media/marker-${idx}.png`, mimeType: 'image/png',
          },
        })),
      })],
    });
    expect(collectChartMarkerImageFills(model(256))).toHaveLength(256);
    expect(collectChartMarkerImageFills(model(257))).toEqual([]);
    expect(collectChartMarkerImageFillsForCharts([model(128), {
      ...model(129),
      series: [series({
        values: Array.from({ length: 129 }, () => 1), showMarker: false,
        dataPointOverrides: Array.from({ length: 129 }, (_, idx) => ({
          idx, markerSymbol: 'picture' as const,
          markerFillPaint: {
            fillType: 'image' as const, stretch: true,
            imagePath: `xl/media/second-${idx}.png`, mimeType: 'image/png',
          },
        })),
      })],
    }])).toEqual([]);
    expect(collectChartMarkerImageFillsForCharts([model(257), model(1)])).toEqual([]);
    expect(collectChartMarkerImageFillsForCharts([model(1), model(257)])).toEqual([]);
  });

  it('does not fetch or charge image fills for stroke-only x/plus symbols', () => {
    const strokeOnly = baseModel({
      chartType: 'line', categories: Array.from({ length: 257 }, (_, index) => String(index)),
      series: [series({
        values: Array.from({ length: 257 }, () => 1), showMarker: false,
        dataPointOverrides: Array.from({ length: 257 }, (_, idx) => ({
          idx, markerSymbol: idx % 2 === 0 ? 'x' : 'plus',
          markerFillPaint: {
            fillType: 'image' as const, stretch: true,
            imagePath: `xl/media/unused-${idx}.png`, mimeType: 'image/png',
          },
        })),
      })],
    });
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/used.png', mimeType: 'image/png',
    };
    const pictureChart = baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1], showMarker: true, markerSymbol: 'picture', markerFillPaint: picture,
      })],
    });
    expect(collectChartMarkerImageFills(strokeOnly)).toEqual([]);
    expect(collectChartMarkerImageFillsForCharts([strokeOnly, pictureChart])).toEqual([picture]);
    const bitmap = { width: 1, height: 1 } as unknown as CanvasImageSource;
    expect(classicMarkerPaintWorkCount(strokeOnly, () => bitmap, 1, RECT)).toBe(0);
  });

  it('uses a radar picture marker for the plot and compound legend key', () => {
    const rec = recordingCtx();
    const bitmap = { width: 24, height: 24 } as unknown as CanvasImageSource;
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/radar.png', mimeType: 'image/png',
    };
    const model = baseModel({
      chartType: 'radar', radarStyle: 'marker', categories: ['A', 'B', 'C'],
      showLegend: true,
      series: [series({
        values: [1, 2, 3], showMarker: true, markerSymbol: 'picture', markerFillPaint: picture,
      })],
    });
    expect(collectChartMarkerImageFills(model)).toEqual([picture]);
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    // 3 plot points + one compound line/marker legend key.
    expect(rec.drawImages).toHaveLength(4);
  });

  it('keeps a direct ChartEx box series color above a linked picture marker', () => {
    const picture = {
      fillType: 'image' as const, stretch: true,
      imagePath: 'xl/media/linked-box.png', mimeType: 'image/png',
    };
    const model = baseModel({
      chartType: 'boxWhisker',
      chartexDataPointMarkerStyle: { fillPaints: [picture], fillPaintAuthored: true },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'Direct', color: 'FF0000', valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: true,
          showNonoutliers: true, quartileMethod: 'exclusive',
        }],
      },
    });
    expect(collectChartMarkerImageFills(model)).toEqual([]);
    const rec = recordingCtx();
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () =>
      ({ width: 16, height: 16 }) as unknown as CanvasImageSource);
    expect(rec.drawImages).toHaveLength(0);
  });

  it.each([
    { seriesType: 'line', markerSymbol: 'square' },
    { seriesType: 'area', markerSymbol: 'circle' },
  ] as const)(
    'uses the shared marker path for a $seriesType overlay in a bar combo',
    ({ seriesType, markerSymbol }) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar', categories: ['A', 'B'],
        catAxisHidden: true, valAxisHidden: true,
        series: [
          series({ name: 'Bars', values: [2, 4], seriesType: 'bar' }),
          series({
            name: 'Overlay', values: [3, 5], seriesType,
            showMarker: true, markerSymbol, markerSize: 12,
            markerFillPaint: {
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: [
                { position: 0, color: '112233' },
                { position: 1, color: 'DDEEFF' },
              ],
            },
          }),
        ],
      }), RECT, 1);
      expect(rec.gradients).toHaveLength(2);
      if (markerSymbol === 'square') {
        expect(rec.rects.filter(rect => rect.w === 12 && rect.h === 12)).toHaveLength(2);
      }
    },
  );

  it.each(['line', 'area'] as const)(
    'renders picture markers for a $seriesType overlay in a bar combo',
    seriesType => {
      const rec = recordingCtx();
      const bitmap = { width: 24, height: 24 } as unknown as CanvasImageSource;
      renderChartCore(rec.ctx, baseModel({
        chartType: 'clusteredBar', categories: ['A', 'B'],
        catAxisHidden: true, valAxisHidden: true,
        series: [
          series({ name: 'Bars', values: [2, 4], seriesType: 'bar' }),
          series({
            name: 'Overlay', values: [3, 5], seriesType,
            showMarker: true, markerSymbol: 'picture', markerSize: 12,
            markerFillPaint: {
              fillType: 'image', stretch: true, imagePath: 'xl/media/combo-marker.png', mimeType: 'image/png',
            },
          }),
        ],
      }), RECT, 1, 0, testThreeD, undefined, () => bitmap);
      expect(rec.drawImages).toHaveLength(2);
      expect(rec.drawImages.every(call => call[0] === bitmap)).toBe(true);
    },
  );

  it('uses linked structured dataPointMarker fill only when direct paint is omitted', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'circle', lineHidden: true,
      })],
      chartStyleRoles: {
        dataPointMarker: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 90,
            stops: [
              { position: 0, color: '010203' },
              { position: 1, color: 'FDFEFF' },
            ],
          }],
        },
      },
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toEqual([
      { position: 0, color: 'rgba(1,2,3,1)' },
      { position: 1, color: 'rgba(253,254,255,1)' },
    ]);
  });

  it('does not replace authored unresolved marker paint with linked style paint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'circle', lineHidden: true,
        markerFillPaintAuthored: true,
      })],
      chartStyleRoles: {
        dataPointMarker: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 90,
            stops: [
              { position: 0, color: '010203' },
              { position: 1, color: 'FDFEFF' },
            ],
          }],
        },
      },
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(0);
  });

  it('does not replace an unsupported linked marker fill with automatic color', () => {
    for (const directProvenance of [undefined, false]) {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line', categories: ['A'],
        catAxisHidden: true, valAxisHidden: true,
        series: [series({
          values: [3], showMarker: true, markerSymbol: 'circle', lineHidden: true,
          markerFillPaintAuthored: directProvenance,
        })],
        chartStyleRoles: {
          dataPointMarker: { fillPaintAuthored: true },
        },
      }), RECT, 1);
      expect(rec.gradients).toHaveLength(0);
      expect(rec.paintEvents.some(event =>
        event.kind === 'fill' && event.fillStyle === '#4472C4'
      )).toBe(false);
      expect(rec.paintEvents).toContainEqual({ kind: 'fill', fillStyle: '#00000000' });
    }
  });

  it('fails closed when an unresolved series marker paint omits the symbol', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, lineHidden: true,
        markerFillPaintAuthored: true,
      })],
    }), RECT, 1);
    expect(rec.paintEvents.some(event =>
      event.kind === 'fill' && event.fillStyle === '#4472C4'
    )).toBe(false);
    expect(rec.paintEvents).toContainEqual({ kind: 'fill', fillStyle: '#00000000' });
  });

  it('keeps point-zero formatting out of the series legend marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'], showLegend: true,
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        name: 'Series', values: [3], showMarker: true, markerSymbol: 'circle',
        lineHidden: true, dataPointColors: ['FF0000'],
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'DDEEFF' },
          ],
        },
      })],
    }), RECT, 1);
    // Point 0 is red, while the legend represents the series-level gradient.
    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(221,238,255,1)' },
    ]);
  });

  it('allows a direct point marker to override a disabled series marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        values: [3, 5], showMarker: false, markerSymbol: 'none', lineHidden: true,
        dataPointOverrides: [{
          idx: 1, markerSymbol: 'circle',
          markerFillPaint: {
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          },
        }],
      })],
    }), RECT, 1);
    expect(rec.gradients).toHaveLength(1);
  });

  it('rejects marker gradients beyond the bounded Canvas stop budget', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [3], showMarker: true, markerSymbol: 'circle',
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: Array.from({ length: 4097 }, (_, index) => ({
            position: index / 4096,
            color: '112233',
          })),
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
  });

  it('charges structured marker paint only for visible sparse points', () => {
    const rec = recordingCtx();
    const count = 10_000;
    const values: Array<number | null> = new Array(count).fill(null);
    values[count - 1] = 3;
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: Array.from({ length: count }, (_, index) => String(index)),
      catAxisHidden: true, valAxisHidden: true,
      series: [series({
        values, showMarker: true, markerSymbol: 'circle', lineHidden: true,
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: Array.from({ length: 4096 }, (_, index) => ({
            position: index / 4095,
            color: '112233',
          })),
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(false);
    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toHaveLength(4096);
  });

  it.each([
    { chartType: 'scatter', scatterStyle: 'lineNoMarker' },
    { chartType: 'radar', radarStyle: 'filled' },
  ] as const)(
    'does not charge suppressed $chartType markers to the paint budget',
    ({ chartType, ...style }) => {
      const rec = recordingCtx();
      const count = 257;
      renderChart(rec.ctx, baseModel({
        chartType,
        ...style,
        categories: Array.from({ length: count }, (_, index) => String(index + 1)),
        catAxisHidden: true, valAxisHidden: true,
        series: [series({
          values: Array.from({ length: count }, (_, index) => index + 1),
          showMarker: true, markerSymbol: 'circle',
          markerFillPaint: {
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: Array.from({ length: 4096 }, (_, index) => ({
              position: index / 4095,
              color: '112233',
            })),
          },
        })],
      }), RECT, 1);
      expect(rec.texts.some(text => text.text === '(too many data points)')).toBe(false);
      expect(rec.gradients).toHaveLength(0);
    },
  );

  it('uses linked marker layout only when a classic marker omits symbol and size', () => {
    const linked = markerRecordingCtx();
    renderChart(linked.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({ values: [3, 5], showMarker: true, lineHidden: true })],
      chartStyleMarkerSymbol: 'square',
      chartStyleMarkerSizePt: 12,
    }), RECT, 1);
    expect(linked.fillRects).toHaveLength(2);
    expect(linked.fillRects.every(rect => rect.w === 12 && rect.h === 12)).toBe(true);

    const direct = markerRecordingCtx();
    renderChart(direct.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({
        values: [3], showMarker: true, lineHidden: true,
        markerSymbol: 'square', markerSize: 4,
      })],
      chartStyleMarkerSymbol: 'diamond',
      chartStyleMarkerSizePt: 12,
    }), RECT, 1);
    expect(direct.fillRects).toHaveLength(1);
    expect(direct.fillRects[0]).toMatchObject({ w: 4, h: 4 });
  });

  it('keeps the marker role off bubble points, which have no classic CT_Marker', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1'],
      catAxisHidden: true,
      valAxisHidden: true,
      series: [series({ values: [3], categories: ['1'], bubbleSizes: [10] })],
      chartStyleRoles: {
        dataPointMarker: { fillColors: ['AABBCC'], lineColors: ['CCBBAA'] },
      },
    }), RECT, 1);
    expect(rec.fills).not.toContain('#AABBCC');
    expect(rec.strokes.some(stroke => stroke.strokeStyle === '#CCBBAA')).toBe(false);
  });

  it('does not revive a noFill series line in the legend from its width or dash', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      showLegend: true,
      legendPos: 'r',
      catAxisHidden: true,
      valAxisHidden: true,
      catAxisLineHidden: true,
      valAxisLineHidden: true,
      series: [series({
        name: 'Hidden stroke',
        values: [3, 5],
        lineHidden: true,
        lineWidthEmu: 25400,
        chartexStyle: {
          lineHidden: true,
          lineWidthEmu: 25400,
          lineDash: 'dash',
          lineCap: 'rnd',
          lineJoin: 'bevel',
        },
      })],
    }), RECT, 1);

    expect(rec.strokes).toHaveLength(0);
  });
});

describe('CH9 — stacked-area markers/labels sit on the fill\'s band top (§21.2.2.32)', () => {
  // CT_AreaChart's ordered `<c:ser>` sequence is also the stacking order:
  // series 0 sits on the category axis and each later series stacks above it.
  // Therefore band si's top is the forward cumulative Σ_{k=0..si}.
  it('a 2-series stacked area places each marker on the forward-cumulative band top', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A'],
      series: [
        series({ name: 'S0', values: [10], showMarker: true }),
        series({ name: 'S1', values: [40], showMarker: true }),
      ],
    }), RECT, 1);
    // One marker arc per series (single category).
    expect(rec.arcs.length).toBe(2);
    const ys = rec.arcs.map(a => a.y).sort((a, b) => a - b);
    // S0 is the bottom band (top=10); S1 is above it (top=10+40=50).
    const [higherY, lowerY] = ys; // higherY = smaller number = higher on screen
    expect(higherY).toBeLessThan(lowerY);
    const s0Y = rec.arcs[0].y;
    const s1Y = rec.arcs[1].y;
    expect(s0Y).toBeGreaterThan(s1Y);
  });

  it('a 3-series stacked area orders markers by forward-cumulative band top', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea',
      categories: ['A'],
      series: [
        series({ name: 'S0', values: [5], showMarker: true }),
        series({ name: 'S1', values: [15], showMarker: true }),
        series({ name: 'S2', values: [30], showMarker: true }),
      ],
    }), RECT, 1);
    expect(rec.arcs.length).toBe(3);
    // Forward-cumulative band tops: S0=5, S1=20, S2=50.
    const [s0Y, s1Y, s2Y] = rec.arcs.map(a => a.y);
    expect(s0Y).toBeGreaterThan(s1Y);
    expect(s1Y).toBeGreaterThan(s2Y);
  });
});

describe('CH9 — line/area draw per-series error bars (§21.2.2.20)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: a series with errBars strokes a vertical bar around each point`, () => {
      const withBars = pathRecordingCtx();
      renderChart(withBars.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({
          name: 'S',
          values: [10, 20, 15],
          errBars: [{ dir: 'y', barType: 'both', plus: [2, 2, 2], minus: [2, 2, 2], noEndCap: false }],
        })],
      }), RECT, 1);
      const without = pathRecordingCtx();
      renderChart(without.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'S', values: [10, 20, 15] })],
      }), RECT, 1);
      // Error bars add vertical segments (constant x, varying y) — 2-vertex
      // "bar" segments the plain plot never emits. Count vertical 2-point segs.
      const verticalSegs = (segs: Array<Array<{ x: number; y: number }>>): number =>
        segs.filter(s => s.length === 2 && Math.round(s[0].x) === Math.round(s[1].x)
          && Math.round(s[0].y) !== Math.round(s[1].y)).length;
      expect(verticalSegs(withBars.segments)).toBeGreaterThan(verticalSegs(without.segments));
    });
  }

  it('uses the linked errorBar role behind direct error-bar properties', () => {
    const rec = segRecordingCtx();
    const model = baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      valAxisMajorGridlines: false,
      series: [series({
        values: [10, 20],
        errBars: [{
          dir: 'y', barType: 'both', plus: [2, 2], minus: [2, 2], noEndCap: true,
        }],
      })],
      chartStyleRoles: {
        errorBar: { lineColors: ['AABBCC'], lineWidthEmu: 19050 },
      },
    });
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.segs.filter(segment => segment.ss === '#AABBCC')).toHaveLength(4);

    model.series[0].errBars![0].color = '112233';
    const direct = segRecordingCtx();
    renderChart(direct.ctx, model, RECT, 1);
    expect(direct.segs.filter(segment => segment.ss === '#112233')).toHaveLength(4);
    expect(direct.segs.some(segment => segment.ss === '#AABBCC')).toBe(false);
  });

  it('honors direct and linked no-fill error-bar strokes', () => {
    const make = (hidden: boolean | undefined, roleHidden: boolean): ChartModel => baseModel({
      chartType: 'line', categories: ['A'], valAxisMajorGridlines: false,
      series: [series({
        values: [10],
        errBars: [{
          dir: 'y', barType: 'plus', plus: [2], minus: [null], noEndCap: true,
          hidden,
        }],
      })],
      chartStyleRoles: { errorBar: { lineColors: ['AABBCC'], lineHidden: roleHidden } },
    });
    for (const model of [make(true, false), make(undefined, true)]) {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, model, RECT, 1);
      expect(rec.segs.some(segment => segment.ss === '#AABBCC')).toBe(false);
    }
  });
});

describe('CH9 — scatter error-bar cap geometry (§21.2.2.20)', () => {
  it('keeps an x-error-bar end cap within an overlaid endpoint marker diameter', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['0.15'],
      series: [series({
        name: 'Start',
        categories: ['0.15'],
        values: [1],
        markerSymbol: 'circle',
        markerSize: 10,
        errBars: [{
          dir: 'x', barType: 'plus', plus: [0.3], minus: [null], noEndCap: false,
          lineWidthEmu: 85725,
        }],
      })],
      valMin: 0,
      valMax: 2,
    }), RECT, 1);
    const marker = rec.arcs.find(a => a.r > 0);
    expect(marker).toBeDefined();
    const cap = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].x - segment[1].x) < 0.001
      && Math.abs((segment[0].y + segment[1].y) / 2 - (marker as ArcCall).y) < 0.001
      && segment[0].x > (marker as ArcCall).x,
    );
    expect(cap).toBeDefined();
    expect(Math.abs((cap as Array<{ x: number; y: number }>)[1].y - (cap as Array<{ x: number; y: number }>)[0].y))
      .toBeLessThanOrEqual((marker as ArcCall).r * 2);
  });
});

describe('CH9 — scatter axis crossing and tick-label position (§21.2.2.207)', () => {
  it('places nextTo labels beyond authored outward tick marks on both numeric axes', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['0', '1'],
      series: [series({ values: [0, 1], showMarker: false })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
      catAxisMajorUnit: 1,
      valAxisMajorUnit: 1,
      catAxisMajorTickMark: 'out',
      valAxisMajorTickMark: 'out',
      catAxisTickLabelPos: 'nextTo',
      valAxisTickLabelPos: 'nextTo',
      catAxisMajorGridlines: false,
      valAxisMajorGridlines: false,
    }), RECT, 1);

    const xZero = rec.texts.find(text =>
      text.text === '0' && text.align === 'center' && text.baseline === 'top');
    const yZero = rec.texts.find(text =>
      text.text === '0' && text.align === 'right' && text.baseline === 'middle');
    expect(xZero).toBeDefined();
    expect(yZero).toBeDefined();

    const xTick = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].x - (xZero as TextCall).x) < 0.001
      && Math.abs(segment[0].x - segment[1].x) < 0.001
      && Math.abs(Math.abs(segment[1].y - segment[0].y) - 6) < 0.001);
    const yTick = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].y - (yZero as TextCall).y) < 0.001
      && Math.abs(segment[0].y - segment[1].y) < 0.001
      && Math.abs(Math.abs(segment[1].x - segment[0].x) - 6) < 0.001);
    expect(xTick).toBeDefined();
    expect(yTick).toBeDefined();

    // Label anchors are measured from the outside endpoint of the tick, not
    // from the axis rule. This keeps the 6pt Office major tick out of the
    // adjacent glyphs at the lower-left crossing.
    expect((xZero as TextCall).y).toBeGreaterThan(Math.max(xTick![0].y, xTick![1].y));
    expect((yZero as TextCall).x).toBeLessThan(Math.min(yTick![0].x, yTick![1].x));
  });

  it('crosses both numeric axes at zero while low tick labels stay on the plot edges', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['-1', '1'],
      series: [series({ values: [-1, 1], showMarker: false })],
      catAxisMin: -1,
      catAxisMax: 1,
      valMin: -1,
      valMax: 1,
      catAxisCrosses: 'autoZero',
      valAxisCrosses: 'autoZero',
      catAxisTickLabelPos: 'low',
      valAxisTickLabelPos: 'low',
      catAxisMajorGridlines: false,
      valAxisMajorGridlines: false,
    }), RECT, 1);

    const verticalAxis = rec.segments
      .filter(segment => segment.length === 2 && Math.abs(segment[0].x - segment[1].x) < 0.001)
      .sort((a, b) => Math.abs(b[1].y - b[0].y) - Math.abs(a[1].y - a[0].y))[0];
    expect(verticalAxis).toBeDefined();
    const horizontalAxis = rec.segments
      .filter(segment => segment.length === 2 && Math.abs(segment[0].y - segment[1].y) < 0.001)
      .sort((a, b) => Math.abs(b[1].x - b[0].x) - Math.abs(a[1].x - a[0].x))[0];
    expect(horizontalAxis).toBeDefined();

    const verticalX = (verticalAxis[0].x + verticalAxis[1].x) / 2;
    const horizontalY = (horizontalAxis[0].y + horizontalAxis[1].y) / 2;
    expect(verticalX).toBeGreaterThan(horizontalAxis[0].x);
    expect(verticalX).toBeLessThan(horizontalAxis[1].x);
    expect(horizontalY).toBeGreaterThan(verticalAxis[0].y);
    expect(horizontalY).toBeLessThan(verticalAxis[1].y);

    const xLabels = rec.texts.filter(text => text.align === 'center' && text.baseline === 'top');
    expect(xLabels.length).toBeGreaterThan(0);
    expect(xLabels.every(text => text.y > Math.max(verticalAxis[0].y, verticalAxis[1].y))).toBe(true);
    const yLabels = rec.texts.filter(text => text.align === 'right' && text.baseline === 'middle');
    expect(yLabels.length).toBeGreaterThan(0);
    expect(yLabels.every(text => text.x < Math.min(horizontalAxis[0].x, horizontalAxis[1].x))).toBe(true);
  });

  it('maps a second scatter group through independent top-X and right-Y axes', () => {
    const rec = markerRecordingCtx();
    const axis = (max: number, divisor: number) => ({
      min: 0,
      max,
      title: null,
      hidden: false,
      formatCode: '0',
      displayUnits: { divisor, builtInUnit: null, label: null },
      lineHidden: false,
      majorTickMark: 'out',
      majorUnit: max / 2,
    });
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['1'],
      series: [
        series({ name: 'Primary', values: [1], markerSize: 6 }),
        series({
          name: 'Secondary',
          categories: ['1000'],
          values: [10],
          markerSize: 6,
          useSecondaryAxis: true,
        }),
      ],
      catAxisMin: 0,
      catAxisMax: 10,
      valMin: 0,
      valMax: 10,
      secondaryCatAxis: axis(2000, 1000),
      secondaryValAxis: axis(20, 10),
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(2);
    expect(rec.arcs[1].x).toBeGreaterThan(rec.arcs[0].x);
    expect(rec.arcs[1].y).toBeLessThan(rec.arcs[0].y);
    expect(rec.texts.some(text => text.text === '2' && text.baseline === 'bottom')).toBe(true);
    expect(rec.texts.some(text => text.text === '2' && text.align === 'left')).toBe(true);
  });
});

describe('CH9 — bubble scale and numeric-X trendlines', () => {
  it('uses the parsed per-point palette for a single bubble series', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        values: [1, 2, 3],
        bubbleSizes: [100, 400, 900],
        dataPointColors: ['4472C4', 'ED7D31', 'A5A5A5'],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);

    expect(rec.arcs.map(arc => arc.fillStyle)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
  });

  it('uses the series bubble3D value for every point as current Excel does', () => {
    const seriesTrue = recordingCtx();
    renderChart(seriesTrue.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        values: [1, 2, 3], bubbleSizes: [100, 100, 100],
        bubble3DGroupDefault: false,
        bubble3D: true,
        dataPointOverrides: [
          { idx: 0, bubble3D: false },
          { idx: 1, bubble3D: true },
        ],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);
    expect(seriesTrue.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(9);

    const seriesFalse = recordingCtx();
    renderChart(seriesFalse.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        values: [1, 2, 3], bubbleSizes: [100, 100, 100],
        bubble3DGroupDefault: true,
        bubble3D: false,
        dataPointOverrides: [
          { idx: 0, bubble3D: true },
          { idx: 2, bubble3D: false },
        ],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);
    expect(seriesFalse.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(0);

    const groupFallback = recordingCtx();
    renderChart(groupFallback.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1'],
      series: [series({
        values: [1], bubbleSizes: [100], bubble3DGroupDefault: true,
        dataPointOverrides: [{ idx: 0, bubble3D: false }],
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);
    expect(groupFallback.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(3);
  });

  it('paints bounded diffuse, shade, and reflected-light material layers in bubble-local coordinates', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble', categories: ['1'],
      series: [series({
        values: [1], bubbleSizes: [100], bubble3D: true,
        chartexStyle: {
          fillColors: ['4472C4'], fillPaintAuthored: true,
        },
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);

    const materials = rec.gradients.filter(gradient => gradient.kind === 'radial');
    expect(materials).toHaveLength(3);
    expect(materials.map(gradient => gradient.stops.length)).toEqual([4, 5, 6]);
    const bubble = rec.arcs.at(-1)!;
    const size = bubble.r * 2;
    const normalizedGradient = (gradient: (typeof materials)[number]) => {
      const [x0, y0, innerRadius, x1, y1, outerRadius] = gradient.args;
      return {
        x0: (x0 - (bubble.x - bubble.r)) / size,
        y0: (y0 - (bubble.y - bubble.r)) / size,
        innerRadius: innerRadius / size,
        x1: (x1 - (bubble.x - bubble.r)) / size,
        y1: (y1 - (bubble.y - bubble.r)) / size,
        outerRadius: outerRadius / size,
      };
    };
    expect(normalizedGradient(materials[0])).toMatchObject({
      x0: expect.closeTo(0.42, 2), y0: expect.closeTo(0.33, 2),
      innerRadius: 0, x1: expect.closeTo(0.42, 2), y1: expect.closeTo(0.33, 2),
      outerRadius: expect.closeTo(0.55, 2),
    });
    expect(normalizedGradient(materials[1])).toMatchObject({
      x0: expect.closeTo(0.42, 2), y0: expect.closeTo(0.32, 2),
      innerRadius: 0, x1: expect.closeTo(0.42, 2), y1: expect.closeTo(0.32, 2),
      outerRadius: expect.closeTo(0.78, 2),
    });
    expect(normalizedGradient(materials[2])).toMatchObject({
      x0: expect.closeTo(0.30, 2), y0: expect.closeTo(0.05, 2),
      innerRadius: 0, x1: expect.closeTo(0.30, 2), y1: expect.closeTo(0.05, 2),
      outerRadius: expect.closeTo(1, 2),
    });
    expect(materials[0].stops.some(stop => stop.color === 'rgba(255,255,255,0.72)'))
      .toBe(true);
    expect(materials[1].stops.some(stop => stop.color === 'rgba(0,0,0,0.62)'))
      .toBe(true);
    expect(materials[2].stops.some(stop => stop.color === 'rgba(255,255,255,0.28)'))
      .toBe(true);
    expect(materials.flatMap(material => material.stops).every(stop =>
      /^rgba\((?:255,255,255|0,0,0),/.test(stop.color)
    )).toBe(true);
    expect(rec.compositeModes.filter(mode => mode === 'source-atop')).toHaveLength(3);
  });

  it('keeps noFill and outline independent from the bubble3D material', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble', categories: ['1'],
      series: [series({
        values: [1], bubbleSizes: [100], bubble3D: true,
        chartexStyle: {
          fillHidden: true, fillPaintAuthored: true,
          lineColors: ['FF0000'], lineWidthEmu: 25_400,
          linePaintAuthored: true,
        },
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);

    expect(rec.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(0);
    expect(rec.strokeDetails.some(stroke =>
      stroke.strokeStyle === 'rgba(255,0,0,1)' && stroke.lineWidth === 2
    )).toBe(true);
  });

  it('always inverts visible negative bubble fill while retaining point outline', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble', showNegativeBubbles: true, categories: ['1'],
      series: [
        series({
          values: [1], bubbleSizes: [-100], bubble3D: false, invertIfNegative: false,
          dataPointOverrides: [{
            idx: 0, color: 'FF0000', lineColor: '7F6000', lineWidthEmu: 25_400,
          }],
        }),
        series({
          values: [1], bubbleSizes: [-100], bubble3D: true, invertIfNegative: true,
          dataPointOverrides: [{
            idx: 0, color: 'FF0000', lineColor: '7F6000', lineWidthEmu: 25_400,
          }],
        }),
      ],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);

    expect(rec.paintEvents.some(event =>
      event.kind === 'fill' && event.fillStyle === '#FF0000'
    )).toBe(false);
    expect(rec.paintEvents.some(event =>
      event.kind === 'fill' && event.fillStyle === '#FFFFFF'
    )).toBe(true);
    expect(rec.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(3);
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '#7F6000')).toHaveLength(2);
  });

  it('uses the observed automatic black outline for a negative 3-D bubble', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble', showNegativeBubbles: true, categories: ['1'],
      series: [series({
        values: [1], bubbleSizes: [-100], bubble3D: true,
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);

    expect(rec.paintEvents.some(event =>
      event.kind === 'fill' && event.fillStyle === '#FFFFFF'
    )).toBe(true);
    expect(rec.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(3);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#000000')).toBe(true);

    const authoredNoLine = recordingCtx();
    renderChart(authoredNoLine.ctx, baseModel({
      chartType: 'bubble', showNegativeBubbles: true, categories: ['1'],
      series: [series({
        values: [1], bubbleSizes: [-100], bubble3D: true, lineHidden: true,
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);
    expect(authoredNoLine.strokeDetails.some(stroke => stroke.strokeStyle === '#000000'))
      .toBe(false);
  });

  it('applies bubble3D to the series legend key as well as the plotted bubble', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble', showLegend: true, legendPos: 'r', categories: ['1'],
      series: [series({
        name: '3-D bubbles', values: [1], bubbleSizes: [100], bubble3D: true,
        chartexStyle: { fillColors: ['70AD47'], fillPaintAuthored: true },
      })],
      catAxisMin: 0, catAxisMax: 2, valMin: 0, valMax: 2,
    }), RECT, 1);

    expect(rec.gradients.filter(gradient => gradient.kind === 'radial')).toHaveLength(6);
  });

  it('lists textual bubble x values as point legend entries', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      showLegend: true,
      legendPos: 'r',
      varyColors: false,
      categories: ['Project A', 'Project B', 'Project C', 'Project D'],
      series: [series({
        name: 'Investment vs Profit',
        bubbleXSourceIsString: true,
        values: [15, 35, 10, 60],
        bubbleSizes: [5, 20, 15, 10],
        dataPointColors: ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000'],
      })],
      valMin: 0,
      valMax: 70,
    }), RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining([
      'Project A', 'Project B', 'Project C', 'Project D',
    ]));
    expect(labels).not.toContain('Investment vs Profit');
    expect(rec.arcs.slice(-4).map(arc => arc.fillStyle)).toEqual([
      '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000',
    ]);
    // Office maps a string-backed bubble X source to one-based ordinal
    // positions, yielding the automatic 0..5 axis for four points.
    expect(labels).toContain('5');
  });

  it('composes showBubbleSize labels with point-level visibility overrides', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2'],
      series: [series({
        values: [2, 3],
        bubbleSizes: [876, 987],
        seriesDataLabels: {
          showVal: false,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          showBubbleSize: true,
        },
        dataLabelOverrides: [{ idx: 0, text: '', showBubbleSize: false }],
      })],
      catAxisMin: 0,
      catAxisMax: 3,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);

    expect(rec.texts.some(text => text.text === '876')).toBe(false);
    expect(rec.texts.some(text => text.text === '987')).toBe(true);
  });

  it('keeps series noFill over varyColors while point formatting stays more specific', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        name: 'Hidden fill',
        color: '00000000',
        values: [1, 2, 3],
        bubbleSizes: [100, 400, 900],
        dataPointColors: [null, 'FF0000', '00000000'],
      })],
      showLegend: true,
      legendPos: 'r',
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);

    expect(rec.arcs.slice(0, 3).map(arc => arc.fillStyle))
      .toEqual(['#00000000', '#FF0000', '#00000000']);
    // The one-series legend key uses the same authored transparent series fill;
    // varyColors does not silently revive it with a theme accent.
    expect(rec.arcs.at(-1)?.fillStyle).toBe('#00000000');
  });

  it('keeps direct bubble shape paint above the linked marker role', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1'],
      chartStyleRoles: {
        dataPointMarker: { fillColors: ['FF0000'], fillPaintAuthored: true },
      },
      series: [series({
        values: [0.25, 0.75],
        bubbleSizes: [100, 100],
        chartexStyle: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          fillPaintAuthored: true,
        },
      })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(2);
    expect(rec.gradients.every(gradient => gradient.stops.length === 2)).toBe(true);
  });

  it('uses the linked dataPoint shape role for unauthored bubbles', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1'],
      chartStyleRoles: {
        dataPoint: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          fillPaintAuthored: true,
        },
      },
      series: [series({ values: [0.25, 0.75], bubbleSizes: [100, 100] })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(2);
  });

  it('keeps legacy direct bubble series paint above the linked dataPoint role', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0'],
      chartStyleRoles: {
        dataPoint: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          fillPaintAuthored: true,
          lineColors: ['00FF00'],
          linePaintAuthored: true,
        },
      },
      series: [series({
        color: 'FF0000',
        lineColor: '0000FF',
        values: [0.5],
        bubbleSizes: [100],
      })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(0);
    expect(rec.paintEvents.some(event => event.kind === 'fill' && event.fillStyle === '#FF0000'))
      .toBe(true);
    expect(rec.paintEvents.some(event => event.kind === 'stroke' && event.strokeStyle === '#0000FF'))
      .toBe(true);
  });

  it('keeps point bubble noFill above direct series structured paint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1'],
      series: [series({
        values: [0.25, 0.75],
        bubbleSizes: [100, 100],
        chartexStyle: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          fillPaintAuthored: true,
        },
        dataPointOverrides: [{
          idx: 0,
          chartexStyle: { fillHidden: true, fillPaintAuthored: true },
        }],
      })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.paintEvents.filter(event => event.kind === 'fill')).toHaveLength(1);
  });

  it('keeps direct bubble outline paint above the linked marker line', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1'],
      chartStyleRoles: {
        dataPointMarker: { lineColors: ['FF0000'], linePaintAuthored: true },
      },
      series: [series({
        values: [0.25, 0.75],
        bubbleSizes: [100, 100],
        chartexStyle: {
          linePaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          linePaintAuthored: true,
          lineWidthEmu: 12700,
        },
        dataPointOverrides: [{
          idx: 0,
          chartexStyle: { lineHidden: true, linePaintAuthored: true },
        }],
      })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
  });

  it('merges point outline geometry over the series bubble paint property by property', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0'],
      chartStyleRoles: {
        dataPoint: { lineColors: ['FF0000'], linePaintAuthored: true },
      },
      series: [series({
        values: [0.5],
        bubbleSizes: [100],
        chartexStyle: {
          linePaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          }],
          linePaintAuthored: true,
        },
        dataPointOverrides: [{
          idx: 0,
          chartexStyle: {
            lineWidthEmu: 25_400,
            lineDash: 'dash',
            lineDashAuthored: true,
            lineCap: 'rnd',
            lineJoin: 'bevel',
          },
        }],
      })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    const bubbleStroke = rec.strokeDetails.find(event =>
      event.strokeStyle === '[object Object]'
    );
    expect(bubbleStroke).toMatchObject({
      lineWidth: 2, cap: 'round', join: 'bevel',
    });
    expect(bubbleStroke?.dash.length).toBeGreaterThan(0);
  });

  it('uses one automatic numeric-axis density for equal X and Y bubble ranges', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({ values: [1, 2, 3], bubbleSizes: [100, 400, 900] })],
    }), { x: 0, y: 0, w: 600, h: 360 }, 1);

    // Both automatic numeric axes use 0..3.5 at 0.5 intervals. Previously the
    // horizontal axis ended at 3.5 while the vertical axis independently chose
    // 0..4, despite having the same source range.
    expect(rec.texts.filter(text => text.text === '3.5')).toHaveLength(2);
    expect(rec.texts.some(text => text.text === '4')).toBe(false);
  });

  it('applies bubbleScale to the default maximum bubble diameter', () => {
    const render = (bubbleScale: number) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        bubbleScale,
        categories: ['0', '1'],
        series: [series({ values: [0, 1], bubbleSizes: [25, 100] })],
        catAxisMin: 0,
        catAxisMax: 1,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return Math.max(...rec.arcs.map(arc => arc.r));
    };
    expect(render(100) / render(50)).toBeCloseTo(1.75, 5);
    expect(render(200) / render(100)).toBeCloseTo(1.6, 5);
  });

  it('normalizes every bubble series against one chart-group maximum', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0'],
      series: [
        series({ values: [0.25], bubbleSizes: [9] }),
        series({ values: [0.75], bubbleSizes: [900] }),
      ],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const radii = rec.arcs.map(arc => arc.r).sort((a, b) => a - b);
    expect(radii).toHaveLength(2);
    expect(radii[0] / radii[1]).toBeCloseTo(0.1, 5);
  });

  it('honors sizeRepresents="w" by making radius proportional to the value', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      bubbleSizeRepresents: 'w',
      categories: ['0', '1'],
      series: [series({ values: [0.25, 0.75], bubbleSizes: [10, 20] })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);

    const radii = rec.arcs.map(arc => arc.r).sort((a, b) => a - b);
    expect(radii).toHaveLength(2);
    expect(radii[1] / radii[0]).toBeCloseTo(2, 5);
  });

  it('excludes non-rendered bubble points from the shared size normalization', () => {
    const render = (withNonRenderedPoints: boolean) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        categories: withNonRenderedPoints ? ['0', '1', 'not-a-number'] : ['0'],
        series: [series({
          values: withNonRenderedPoints ? [0.5, null, 0.75] : [0.5],
          bubbleSizes: withNonRenderedPoints ? [100, 1_000_000, 1_000_000] : [100],
        })],
        catAxisMin: 0,
        catAxisMax: 1,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return rec.arcs;
    };

    const baseline = render(false);
    const sparse = render(true);
    expect(sparse).toHaveLength(1);
    expect(sparse[0].r).toBeCloseTo(baseline[0].r, 5);
  });

  it('does not draw bubbles when the chart scale is zero', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      bubbleScale: 0,
      categories: ['0'],
      series: [series({ values: [0.5], bubbleSizes: [100] })],
      catAxisMin: 0,
      catAxisMax: 1,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(0);
  });

  it('draws only positive bubble sizes and honors per-point marker suppression', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0', '1', '2', '3', '4'],
      series: [series({
        values: [0, 0.25, 0.5, 0.75, 1],
        bubbleSizes: [100, 0, -10, null, 1_000_000],
        dataPointOverrides: [{ idx: 4, markerSymbol: 'none' }],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 1,
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(1);
  });

  it('draws negative bubble sizes by absolute magnitude only when showNegBubbles is true', () => {
    const render = (showNegativeBubbles: boolean | undefined) => {
      const rec = markerRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'bubble',
        showNegativeBubbles,
        categories: ['0', '1', '2'],
        series: [series({ values: [0.25, 0.5, 0.75], bubbleSizes: [-100, 0, 25] })],
        catAxisMin: 0,
        catAxisMax: 2,
        valMin: 0,
        valMax: 1,
      }), RECT, 1);
      return rec.arcs;
    };

    expect(render(undefined)).toHaveLength(1);
    const enabled = render(true);
    expect(enabled).toHaveLength(2);
    expect(enabled[0].r).toBeGreaterThan(enabled[1].r);
  });

  it('fits and extends a scatter trendline in numeric X-axis units', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1', '2', '3'],
      series: [series({
        values: [1, 2, 3],
        bubbleSizes: [1, 1, 1],
        trendLines: [{ trendlineType: 'linear', backward: 0.5, forward: 0.5, lineColor: '000000' }],
      })],
      catAxisMin: 0,
      catAxisMax: 4,
      valMin: 0,
      valMax: 4,
    }), RECT, 1);
    const markerXs = rec.arcs.map(arc => arc.x);
    const diagonal = rec.segments.find(segment =>
      segment.length === 2
      && Math.abs(segment[0].x - segment[1].x) > 1
      && Math.abs(segment[0].y - segment[1].y) > 1,
    );
    expect(diagonal).toBeDefined();
    const trendline = diagonal as Array<{ x: number; y: number }>;
    expect(Math.min(trendline[0].x, trendline[1].x)).toBeLessThan(Math.min(...markerXs));
    expect(Math.max(trendline[0].x, trendline[1].x)).toBeGreaterThan(Math.max(...markerXs));
  });
});

describe('classic data-label legend keys (§21.2.2.179)', () => {
  const baseChart = (): ChartModel => ({
    chartType: 'clusteredBar',
    categories: ['A', 'B'],
    series: [{
      name: 'Series 1',
      values: [10, 20],
      color: '4472C4',
      seriesDataLabels: {
        showVal: true,
        showCatName: false,
        showSerName: false,
        showPercent: false,
        showLegendKey: true,
      },
    }],
    showLegend: false,
  } as ChartModel);

  it('paints the resolved series key beside each column label', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseChart(), { x: 0, y: 0, w: 500, h: 300 }, 1);

    const keys = rec.rects.filter(rect =>
      Math.abs(rect.w - 7) < 0.01 && Math.abs(rect.h - 7) < 0.01 && rect.fs === '#4472C4'
    );
    expect(keys).toHaveLength(2);
    const labels = rec.texts.filter(call =>
      (call.text === '10' || call.text === '20') && call.fillStyle === '#333'
    );
    expect(labels).toHaveLength(2);
    expect(keys.every(key => labels.some(text =>
      text.x > key.x + key.w && text.x - (key.x + key.w) <= 5
    ))).toBe(true);
  });

  it('supports a key-only label and a per-point false override', () => {
    const chart = baseChart();
    const series = chart.series[0];
    series.seriesDataLabels = {
      showVal: false,
      showCatName: false,
      showSerName: false,
      showPercent: false,
      showLegendKey: true,
    };
    series.dataLabelOverrides = [{ idx: 1, text: '', showLegendKey: false }];
    const rec = recordingCtx();
    renderChart(rec.ctx, chart, { x: 0, y: 0, w: 500, h: 300 }, 1);

    expect(rec.rects.filter(rect =>
      Math.abs(rect.w - 7) < 0.01 && Math.abs(rect.h - 7) < 0.01 && rect.fs === '#4472C4'
    )).toHaveLength(1);
  });

  it('uses the effective per-slice color for pie label keys', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      chartType: 'pie',
      categories: ['A', 'B'],
      series: [{
        name: 'Series 1',
        values: [1, 1],
        color: '4472C4',
        dataPointColors: ['FF0000', '00FF00'],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          showLegendKey: true,
          position: 'ctr',
        },
      }],
      showLegend: false,
    } as ChartModel, { x: 0, y: 0, w: 500, h: 300 }, 1);

    const keys = rec.rects.filter(rect => Math.abs(rect.w - 7) < 0.01 && Math.abs(rect.h - 7) < 0.01);
    expect(keys.map(rect => rect.fs)).toEqual(expect.arrayContaining(['#FF0000', '#00FF00']));
  });
});

describe('showDLblsOverMax (§21.2.2.180)', () => {
  const labels = {
    showVal: true,
    showCatName: false,
    showSerName: false,
    showPercent: false,
    fontColor: 'FF00FF',
  };

  const renderedLabelTexts = (chart: ChartModel): string[] => {
    const rec = recordingCtx();
    renderChart(rec.ctx, chart, { x: 0, y: 0, w: 500, h: 300 }, 1);
    return rec.texts
      .filter(call => call.fillStyle === '#FF00FF')
      .map(call => call.text);
  };

  it('suppresses values above the effective maximum unless explicitly enabled', () => {
    const chart = {
      chartType: 'clusteredBar',
      categories: ['inside', 'over'],
      series: [{
        name: 'Series 1', values: [5, 15], color: '4472C4', seriesDataLabels: labels,
      }],
      valMin: 0,
      valMax: 10,
      showLegend: false,
    } as ChartModel;
    expect(renderedLabelTexts(chart)).toEqual(['5']);
    chart.showDataLabelsOverMax = true;
    expect(renderedLabelTexts(chart)).toEqual(['5', '15']);
  });

  it('compares each stacked data-point value, not the cumulative endpoint, to the numeric maximum', () => {
    const stacked = {
      chartType: 'stackedBar',
      categories: ['stack'],
      series: [
        { name: 'Base', values: [8], color: '4472C4', seriesDataLabels: labels },
        { name: 'Top', values: [8], color: 'ED7D31', seriesDataLabels: labels },
      ],
      valMin: 0,
      valMax: 10,
      showLegend: false,
    } as ChartModel;
    expect(renderedLabelTexts(stacked)).toEqual(['8', '8']);

    const roundedPercentTotal = {
      chartType: 'stackedBarPct',
      categories: ['rounded'],
      series: [
        { name: 'Base', values: [0.18], color: '4472C4' },
        { name: 'Middle', values: [0.456], color: 'ED7D31' },
        { name: 'Top', values: [0.365], color: 'A5A5A5', seriesDataLabels: labels },
      ],
      valMin: 0,
      valMax: 1,
      showLegend: false,
    } as ChartModel;
    expect(renderedLabelTexts(roundedPercentTotal)).toEqual(['0.365']);

    stacked.series[1].values = [15];
    expect(renderedLabelTexts(stacked)).toEqual(['8']);

    const negative = {
      chartType: 'clusteredBar',
      categories: ['inside', 'over'],
      series: [{
        name: 'Negative', values: [-10, -2], color: '4472C4', seriesDataLabels: labels,
      }],
      valMin: -12,
      valMax: -5,
      showLegend: false,
    } as ChartModel;
    expect(renderedLabelTexts(negative)).toEqual(['-10']);
  });

  it('uses the owning secondary axis and applies the gate to point-level labels', () => {
    const chart = {
      chartType: 'line',
      categories: ['inside', 'over'],
      series: [{
        name: 'Secondary',
        values: [5, 15],
        color: '4472C4',
        useSecondaryAxis: true,
        seriesDataLabels: {
          ...labels,
          showVal: false,
        },
        dataLabelOverrides: [
          { idx: 0, text: '', showVal: true, fontColor: 'FF00FF' },
          { idx: 1, text: '', showVal: true, fontColor: 'FF00FF' },
        ],
      }],
      secondaryValAxis: { min: 0, max: 10 },
      showLegend: false,
    } as ChartModel;
    expect(renderedLabelTexts(chart)).toEqual(['5']);
    chart.showDataLabelsOverMax = true;
    expect(renderedLabelTexts(chart)).toEqual(['5', '15']);
  });

  it('applies the same resolved maximum to classic 3-D labels', () => {
    const chart = {
      chartType: 'clusteredBar',
      categories: ['inside', 'over'],
      series: [{
        name: 'Series 1', values: [5, 15], color: '4472C4', seriesDataLabels: labels,
      }],
      valMin: 0,
      valMax: 10,
      showLegend: false,
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
    } as ChartModel;
    expect(renderedLabelTexts(chart)).toEqual(['5']);
    chart.showDataLabelsOverMax = true;
    expect(renderedLabelTexts(chart)).toEqual(['5', '15']);
  });
});

describe('CH9 — line/area per-point data labels (§21.2.2.45)', () => {
  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: dataLabelOverrides render custom text at the point, and delete (empty) skips it`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({
          name: 'S',
          values: [3, 5, 4],
          dataLabelOverrides: [
            { idx: 0, text: 'FIRST' },
            { idx: 1, text: '' }, // deleted
            { idx: 2, text: 'THIRD', fontColor: 'FF0000' },
          ],
        })],
      }), RECT, 1);
      const labelTexts = rec.texts.map(t => t.text);
      expect(labelTexts).toContain('FIRST');
      expect(labelTexts).toContain('THIRD');
      // The deleted (empty) label must not appear.
      expect(labelTexts.some(t => t === '')).toBe(false);
    });

    it(`${chartType}: seriesDataLabels showVal renders each point's value`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A', 'B'],
        series: [series({
          name: 'S',
          values: [42, 7],
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
          },
        })],
      }), RECT, 1);
      expect(rec.texts.some(t => t.text === '42')).toBe(true);
      expect(rec.texts.some(t => t.text === '7')).toBe(true);
    });
  }

  for (const chartType of ['stackedLine', 'stackedArea'] as const) {
    it(`${chartType}: showVal uses the source value while the anchor stays cumulative`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        series: [
          series({
            name: 'Lower', values: [10],
            seriesDataLabels: {
              showVal: true, showCatName: false, showSerName: false, showPercent: false,
              fontColor: 'FF0000', position: 'ctr',
            },
          }),
          series({
            name: 'Upper', values: [40],
            seriesDataLabels: {
              showVal: true, showCatName: false, showSerName: false, showPercent: false,
              fontColor: '0000FF', position: 'ctr',
            },
          }),
        ],
      }), RECT, 1);
      expect(rec.texts.find(text => text.fillStyle === '#FF0000')?.text).toBe('10');
      const upper = rec.texts.find(text => text.fillStyle === '#0000FF');
      expect(upper?.text).toBe('40');
      const cumulativeTick = rec.texts.find(text => text.text === '50');
      expect(cumulativeTick).toBeDefined();
      expect(upper?.y).toBeCloseTo(cumulativeTick?.y as number, 6);
    });
  }

  for (const chartType of ['stackedLinePct', 'stackedAreaPct'] as const) {
    it(`${chartType}: showPercent uses the source contribution and authored numFmt`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        series: [
          series({
            values: [1],
            seriesDataLabels: {
              showVal: false, showCatName: false, showSerName: false, showPercent: true,
              formatCode: '0.0%', fontColor: 'FF0000', position: 'ctr',
            },
          }),
          series({ values: [2] }),
        ],
      }), RECT, 1);
      expect(rec.texts.find(text => text.fillStyle === '#FF0000')?.text).toBe('33.3%');
    });
  }

  it('percent-stacked labels use authored composition, separator, and number format', () => {
    const rec = recordingCtx();
    const labels = {
      showVal: false, showCatName: true, showSerName: false, showPercent: true,
      separator: '  ', formatCode: '0.0%', position: 'ctr',
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBarPct', categories: ['A'],
      series: [
        series({ name: 'One', values: [1], seriesDataLabels: labels }),
        series({ name: 'Two', values: [2], seriesDataLabels: labels }),
      ],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === 'A  33.3%')).toBe(true);
    expect(rec.texts.some(text => text.text === '33%')).toBe(false);
  });

  it('per-point text/manual layout/style wins while format, separator, and delete stay indexed', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({
        name: 'S',
        values: [42, 7],
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          formatCode: '0', separator: ' ', position: 'r',
        },
        dataLabelOverrides: [
          {
            idx: 0, text: '', showVal: true, showCatName: true,
            formatCode: '0.0', separator: '|', position: 'l',
            fontColor: 'FF0000', fontBold: true, fontSizeHpt: 1200,
            manualLayout: {
              xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5, w: 0.1, h: 0.1,
            },
          },
          { idx: 1, text: 'HIDDEN', deleted: true },
        ],
      })],
    }), RECT, 1);

    const label = rec.texts.find(text => text.text === 'A|42.0');
    expect(label).toMatchObject({ align: 'center', baseline: 'middle', fillStyle: '#FF0000' });
    expect(label?.font).toContain('bold 12px');
    expect(rec.texts.some(text => text.text === 'HIDDEN' || text.text === 'B 7')).toBe(false);
    expect(rec.clips.some(clip => clip.w > 0 && clip.h > 0 && clip.w < RECT.w / 2)).toBe(true);
  });
});

describe('CH11 — line/area/scatter data labels honor <c:dLblPos> (§21.2.2.48)', () => {
  // drawDataLabelText encodes each position purely through textAlign/textBaseline
  // (+ a directional offset), so the recorded align/baseline of a value label is
  // a faithful witness of the resolved <c:dLblPos>:
  //   r → left/middle   l → right/middle   t → center/bottom
  //   b → center/top     ctr → center/middle
  const expectPos: Record<string, { align: string; baseline: string }> = {
    r: { align: 'left', baseline: 'middle' },
    l: { align: 'right', baseline: 'middle' },
    t: { align: 'center', baseline: 'bottom' },
    b: { align: 'center', baseline: 'top' },
    ctr: { align: 'center', baseline: 'middle' },
  };
  // Find the value label for the single data point (text "42").
  const valLabel = (rec: Recorded): TextCall => {
    const hit = rec.texts.find(t => t.text === '42');
    if (!hit) throw new Error('value label "42" not drawn');
    return hit;
  };

  it('line: the default position is r (right of the point) per PowerPoint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({ name: 'S', values: [42], showMarker: true })],
      showDataLabels: true,       // family-level value dump (legacy path)
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');    // right-of-point → left-aligned text
    expect(lbl.baseline).toBe('middle');
  });

  it('line: seriesDataLabels default position is r when no dLblPos is set', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');
    expect(lbl.baseline).toBe('middle');
  });

  for (const pos of ['t', 'b', 'l', 'r', 'ctr'] as const) {
    it(`line: an explicit <c:dLblPos val="${pos}"> places the label ${pos}`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A'],
        series: [series({
          name: 'S', values: [42], showMarker: true,
          seriesDataLabels: {
            showVal: true, showCatName: false, showSerName: false, showPercent: false,
            position: pos,
          },
        })],
      }), RECT, 1);
      const lbl = valLabel(rec);
      expect(lbl.align).toBe(expectPos[pos].align);
      expect(lbl.baseline).toBe(expectPos[pos].baseline);
    });

    it(`line: a chart-level dataLabelPosition="${pos}" flows to the family value dump`, () => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['A'],
        series: [series({ name: 'S', values: [42], showMarker: true })],
        showDataLabels: true,
        dataLabelPosition: pos,
      }), RECT, 1);
      const lbl = valLabel(rec);
      expect(lbl.align).toBe(expectPos[pos].align);
      expect(lbl.baseline).toBe(expectPos[pos].baseline);
    });
  }

  it('line: a per-point override position beats the series-level position', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          position: 'r',
        },
        dataLabelOverrides: [{ idx: 0, text: '42', position: 't' }],
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('center');  // 't' wins over series 'r'
    expect(lbl.baseline).toBe('bottom');
  });

  it('line: endpoint labels stay outside the plot and clear authored markers', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Start', 'End'],
      catAxisCrossBetween: 'midCat',
      plotAreaManualLayout: {
        layoutTarget: 'inner', xMode: 'edge', yMode: 'edge',
        x: 0.2, y: 0.2, w: 0.6, h: 0.6,
      },
      series: [series({
        name: 'Slope', values: [40, 60], showMarker: true,
        markerSymbol: 'circle', markerSize: 7,
        dataLabelOverrides: [
          { idx: 0, text: 'Left endpoint', position: 'l', fontSizeHpt: 1200 },
          { idx: 1, text: 'Right endpoint', position: 'r', fontSizeHpt: 1200 },
        ],
      })],
    }), RECT, 1);

    const markers = rec.arcs.filter(arc => Math.abs(arc.r - 3.5) < 1e-6);
    expect(markers).toHaveLength(2);
    const left = rec.texts.find(call => call.text === 'Left endpoint');
    const right = rec.texts.find(call => call.text === 'Right endpoint');
    expect(left).toBeDefined();
    expect(right).toBeDefined();
    // fillText x is the text's right edge for `l`, left edge for `r`.
    // Office separates each label from the marker center by marker radius +
    // half an em: 3.5pt + 6pt at the authored 12pt label size.
    expect(left?.x).toBeCloseTo(markers[0].x - 9.5, 4);
    expect(right?.x).toBeCloseTo(markers[1].x + 9.5, 4);
    expect(left?.x).toBeLessThan(RECT.w * 0.2);
    expect(right?.x).toBeGreaterThan(RECT.w * 0.8);
  });

  it('line: custom rich label runs stay inline with per-run size and weight', () => {
    const rich = recordingCtx();
    const model = baseModel({
      chartType: 'line',
      categories: ['Start'],
      series: [series({
        name: 'Employer', values: [36], showMarker: true,
        markerSymbol: 'circle', markerSize: 7,
        dataLabelOverrides: [{
          idx: 0,
          text: 'Employer 36.0%',
          position: 'l',
          fontColor: '111111',
          fontSizeHpt: 1200,
          fontBold: true,
          richRuns: [
            {
              text: 'Employer', fontSizeHpt: 1200, bold: true,
              color: '1696D2', fontFace: 'Lato',
            },
            { text: ' 36.0%', fontSizeHpt: 1100, bold: false, color: '333333' },
          ],
        }],
      })],
    });
    renderChart(rich.ctx, model, RECT, 1);

    const employer = rich.texts.find(text => text.text === 'Employer');
    const value = rich.texts.find(text => text.text === ' 36.0%');
    expect(employer).toMatchObject({ fillStyle: '#1696D2', align: 'left', baseline: 'middle' });
    expect(value).toMatchObject({ fillStyle: '#333333', align: 'left', baseline: 'middle' });
    expect(employer?.font).toContain('bold 12px "Lato"');
    expect(value?.font).toContain('11px sans-serif');
    expect(value?.font).not.toContain('bold ');
    expect(value?.x).toBeCloseTo((employer?.x ?? 0) + (employer?.width ?? 0), 6);

    // The rich line is measured as one object. Its final right edge stays at
    // the exact point-label anchor used by the established flattened path.
    const plain = recordingCtx();
    const plainModel = structuredClone(model);
    delete plainModel.series[0].dataLabelOverrides?.[0].richRuns;
    renderChart(plain.ctx, plainModel, RECT, 1);
    const flattened = plain.texts.find(text => text.text === 'Employer 36.0%');
    expect((value?.x ?? 0) + (value?.width ?? 0)).toBeCloseTo(flattened?.x ?? 0, 6);
  });

  it('line: custom rich labels are bounded to 4096 scalars and four lines', () => {
    const paragraph = 'x'.repeat(1100);
    const text = Array.from({ length: 5 }, () => paragraph).join('\n');
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text, position: 'ctr',
          richRuns: [{ text, fontSizeHpt: 1000 }],
        }],
      })],
    }), RECT, 1);

    const painted = rec.texts.filter(call => /^x+$/.test(call.text));
    expect(painted).toHaveLength(4);
    expect(painted.reduce((count, call) => count + Array.from(call.text).length, 0))
      .toBeLessThanOrEqual(4096);
  });

  it('line: bounds a public rich label before whole-string normalization', () => {
    const guarded = new String('x'.repeat(5000));
    Object.defineProperty(guarded, 'replace', {
      value: () => { throw new Error('must not normalize the unbounded source'); },
    });
    const rec = recordingCtx();
    expect(() => renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'bounded', position: 'ctr',
          richRuns: [{ text: guarded as unknown as string, fontSizeHpt: 1000 }],
        }],
      })],
    }), RECT, 1)).not.toThrow();

    expect(rec.texts.filter(call => /^x+$/.test(call.text)))
      .toHaveLength(1);
    expect(rec.texts.find(call => /^x+$/.test(call.text))?.text)
      .toHaveLength(4096);
  });

  it('line: caps an adversarial public rich-run array before scanning empty runs', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'bounded', position: 'ctr',
          richRuns: [
            ...Array.from({ length: 4096 }, () => ({ text: '' })),
            { text: 'must-not-be-scanned' },
          ],
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.map(call => call.text)).not.toContain('must-not-be-scanned');
  });

  it.each([
    [99, '10px'],
    [100, '1px'],
    [400001, '10px'],
  ] as const)('line: safely resolves public rich-run font size %i', (fontSizeHpt, expectedFont) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'Safe', position: 'ctr', fontSizeHpt: 1000,
          richRuns: [{ text: 'Safe', fontSizeHpt }],
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(call => call.text === 'Safe')?.font).toContain(expectedFont);
  });

  it('line: accepts the ST_TextFontSize upper bound exactly', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'Huge', position: 'ctr', fontSizeHpt: 1000,
          richRuns: [{ text: 'Huge', fontSizeHpt: 400000 }],
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(call => call.text === 'Huge')?.font).toContain('4000px');
  });

  it('scatter: custom rich label runs use the shared bounded point-label path', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter', categories: ['2'],
      series: [series({
        values: [3],
        dataLabelOverrides: [{
          idx: 0, text: 'A 3', position: 'r',
          richRuns: [
            { text: 'A', fontSizeHpt: 1200, bold: true },
            { text: ' 3', fontSizeHpt: 1100, bold: false },
          ],
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(call => call.text === 'A')?.font).toContain('bold 12px');
    expect(rec.texts.find(call => call.text === ' 3')?.font).toContain('11px');
  });

  it('scatter: an automatic right endpoint label may use chart gutter and clears its marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      scatterStyle: 'marker',
      catAxisMin: 0,
      catAxisMax: 1.4,
      valMin: 0,
      valMax: 1,
      plotAreaManualLayout: {
        layoutTarget: 'inner', xMode: 'edge', yMode: 'edge',
        x: 0.2, y: 0.2, w: 0.6, h: 0.6,
      },
      series: [series({
        categories: ['1.32'],
        catFormatCode: '0%',
        values: [0.5],
        showMarker: true,
        markerSymbol: 'circle',
        markerSize: 10,
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
        },
      })],
    }), RECT, 1);

    const marker = rec.arcs.find(arc => Math.abs(arc.r - 5) < 1e-6);
    const label = rec.texts.find(text => text.text === '132%');
    expect(marker).toBeDefined();
    expect(label).toMatchObject({ align: 'left', baseline: 'middle' });
    expect((label as TextCall).x).toBeGreaterThan((marker?.x ?? 0) + (marker?.r ?? 0));
    // The authored inner plot ends at 80% of chart width. Office permits this
    // automatic endpoint label to occupy the surrounding chart-area gutter.
    expect((label as TextCall).x + ((label as TextCall).width ?? 0))
      .toBeGreaterThan(RECT.w * 0.8);
  });

  it('bar-scatter combo: an automatic right label retains marker clearance at the chart edge', () => {
    const rec = recordingCtx();
    const values = [14.5, null, null, 11.5, 10.5, 9.5, 8.5, 7.5,
      null, null, 4.5, 3.5, 2.5, 1.5, 0.5];
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: [
        'Overall', '', 'Group A',
        'Long category item A1', 'Long category item A2',
        'Long category item A3', 'Long category item A4',
        'Long category item A5', '', 'Group B',
        'Long category item B1', 'Long category item B2',
        'Long category item B3', 'Long category item B4',
        'Long category item B5',
      ],
      valMax: 1.4,
      catAxisOrientation: 'maxMin',
      catAxisFontSizeHpt: 1200,
      secondaryCatAxis: {
        min: null, max: null, title: null, hidden: true, lineHidden: true,
        majorTickMark: 'none', formatCode: '0%',
      },
      secondaryValAxis: {
        min: 0, max: 15, title: null, hidden: true, lineHidden: true,
        majorTickMark: 'none',
      },
      series: [
        series({
          values: values.map(value => value == null ? null : 0),
          seriesType: 'bar',
          showMarker: false,
        }),
        series({
          values,
          seriesType: 'scatter',
          useSecondaryAxis: true,
          categories: ['0.15', '', '', '0.83', '0.67', '0.58', '0.58', '0.54',
            '', '', '0.75', '0.54', '0.53', '0.52', '0.49'],
          showMarker: true,
          markerSymbol: 'circle',
          markerSize: 10,
          seriesDataLabels: {
            showVal: false,
            showCatName: true,
            showSerName: false,
            showPercent: false,
            position: 'l',
            fontSizeHpt: 1200,
          },
          catFormatCode: '0%',
        }),
        series({
          values,
          seriesType: 'scatter',
          useSecondaryAxis: true,
          categories: ['0.45', '', '', '1.32', '1.11', '0.99', '0.99', '0.95',
            '', '', '1.21', '0.95', '0.94', '0.92', '0.88'],
          showMarker: true,
          markerSymbol: 'circle',
          markerSize: 10,
          seriesDataLabels: {
            showVal: false,
            showCatName: true,
            showSerName: false,
            showPercent: false,
            fontSizeHpt: 1200,
            textLInsEmu: 38100,
            textRInsEmu: 38100,
            textTInsEmu: 19050,
            textBInsEmu: 19050,
            textBodyAuthored: true,
          },
          catFormatCode: '0%',
        }),
        series({
          values,
          seriesType: 'scatter',
          useSecondaryAxis: true,
          categories: values.map(() => '0'),
          showMarker: false,
          markerSymbol: 'none',
        }),
      ],
      plotGroups: [
        plotGroup('bar', 0, 1, { grouping: 'clustered', barDirection: 'bar' }),
        plotGroup('scatter', 1, 3, {
          categoryAxis: 'secondary', valueAxis: 'secondary', scatterStyle: 'lineMarker',
        }),
      ],
    }), { x: 0, y: 0, w: 960, h: 540 }, 4 / 3);

    const label = rec.texts.find(text => text.text === '132%');
    const marker = rec.arcs
      .filter(arc => Math.abs(arc.y - (label?.y ?? Number.POSITIVE_INFINITY)) < 1e-6)
      .sort((a, b) => b.x - a.x)[0];
    expect(marker).toBeDefined();
    expect(label).toMatchObject({ align: 'left', baseline: 'middle' });
    expect((label as TextCall).x)
      .toBeGreaterThanOrEqual((marker?.x ?? 0) + (marker?.r ?? 0) + 6);
  });

  it.each(['line', 'area', 'scatter'] as const)(
    '%s: rich label run faces resolve major/minor theme references and concrete faces',
    (chartType) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['1'],
        themeMajorFontLatin: 'Major Theme',
        themeMinorFontLatin: 'Minor Theme',
        dataLabelFontFace: '+mn-lt',
        series: [series({
          values: [1],
          dataLabelOverrides: [{
            idx: 0,
            text: 'Major Minor Concrete',
            position: 'ctr',
            richRuns: [
              { text: 'Major', fontFace: '+mj-lt' },
              { text: ' Minor', fontFace: '+mn-lt' },
              { text: ' Concrete', fontFace: 'Direct Face' },
            ],
          }],
        })],
      }), RECT, 1);

      expect(rec.texts.find(call => call.text === 'Major')?.font)
        .toContain('"Major Theme"');
      expect(rec.texts.find(call => call.text === ' Minor')?.font)
        .toContain('"Minor Theme"');
      expect(rec.texts.find(call => call.text === ' Concrete')?.font)
        .toContain('"Direct Face"');
    },
  );

  it('line: all series geometry is painted before the chart-wide label layer', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [
        series({
          name: 'Earlier', values: [40, 60], lineColor: 'FF0000',
          dataLabelOverrides: [{ idx: 0, text: 'Earlier label', position: 'l' }],
        }),
        series({
          name: 'Later', values: [60, 40], lineColor: '0000FF',
          dataLabelOverrides: [{ idx: 0, text: 'Later label', position: 'l' }],
        }),
      ],
    }), RECT, 1);

    const blueSeriesStroke = rec.paintEvents.findIndex(
      event => event.kind === 'stroke' && event.strokeStyle === '#0000FF',
    );
    const firstDataLabel = rec.paintEvents.findIndex(
      event => event.kind === 'text' && event.text.endsWith('label'),
    );
    expect(blueSeriesStroke).toBeGreaterThanOrEqual(0);
    expect(firstDataLabel).toBeGreaterThan(blueSeriesStroke);
  });

  it('area: the default position is ctr (centered on the point) per the areaChart group', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['A'],
      series: [series({
        name: 'S', values: [42], showMarker: true,
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('center');
    expect(lbl.baseline).toBe('middle');
  });

  it('scatter: the default position stays r (unchanged)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['1'],
      series: [series({
        name: 'S', values: [42],
        seriesDataLabels: { showVal: true, showCatName: false, showSerName: false, showPercent: false },
      })],
    }), RECT, 1);
    const lbl = valLabel(rec);
    expect(lbl.align).toBe('left');
    expect(lbl.baseline).toBe('middle');
  });
});

describe('CH9 — line/area smooth splines (§21.2.2.194)', () => {
  it('uses Excel automatic smooth lines for an unformatted marker scatter', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      scatterStyle: 'marker',
      categories: ['1', '2', '3', '4'],
      series: [series({ name: 'S', values: [10, 15, 12, 22], showMarker: true })],
    }), RECT, 1);
    expect(rec.beziers).toBeGreaterThan(0);
  });

  it('keeps filtered automatic scatter topology and colors local to its series', () => {
    const model = baseModel({
      chartType: 'scatter',
      scatterStyle: 'marker',
      plotVisibleOnly: true,
      themeAccentColors: ['156082', 'E97132', '196B24'],
      categories: ['1', '2', '3', '4'],
      series: [
        series({
          values: [10, 20, 30, 40],
          sourceHidden: [false, true, false, false],
        }),
        series({
          color: '7030A0',
          values: [12, 18, 24, 36],
          sourceHidden: [false, true, false, false],
          markerSymbol: 'circle',
        }),
      ],
    });
    const filtered = markerRecordingCtx();
    renderChart(filtered.ctx, model, RECT, 1);

    expect(filtered.strokeStyles).toEqual(expect.arrayContaining(['#E97132', '#196B24']));
    // The directly formatted sibling retains the ordinary marker-scatter
    // spline; the filtered automatic series alone changes to straight,
    // point-colored segments.
    expect(filtered.beziers).toBeGreaterThan(0);

    const visibleOnlyOff = markerRecordingCtx();
    renderChart(visibleOnlyOff.ctx, { ...model, plotVisibleOnly: false }, RECT, 1);
    expect(visibleOnlyOff.strokeStyles).not.toEqual(
      expect.arrayContaining(['#E97132', '#196B24']),
    );
    expect(visibleOnlyOff.beziers).toBeGreaterThan(filtered.beziers);
  });

  for (const chartType of ['line', 'area'] as const) {
    it(`${chartType}: smooth series draws a bezier spline; non-smooth draws straight segments`, () => {
      const smooth = markerRecordingCtx();
      renderChart(smooth.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C', 'D'],
        series: [series({ name: 'S', values: [3, 5, 4, 6], smooth: true })],
      }), RECT, 1);
      const straight = markerRecordingCtx();
      renderChart(straight.ctx, baseModel({
        chartType,
        categories: ['A', 'B', 'C', 'D'],
        series: [series({ name: 'S', values: [3, 5, 4, 6] })],
      }), RECT, 1);
      expect(smooth.beziers).toBeGreaterThan(0);
      expect(straight.beziers).toBe(0);
    });
  }

  it('bar + line combo applies the line-series smooth flag in the overlay path', () => {
    const smooth = markerRecordingCtx();
    renderChart(smooth.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      series: [
        series({ name: 'Bars', values: [30, 45, 40, 60], seriesType: 'bar' }),
        series({ name: 'Rate', values: [3, 5, 4, 6], seriesType: 'line', smooth: true }),
      ],
    }), RECT, 1);

    const straight = markerRecordingCtx();
    renderChart(straight.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      series: [
        series({ name: 'Bars', values: [30, 45, 40, 60], seriesType: 'bar' }),
        series({ name: 'Rate', values: [3, 5, 4, 6], seriesType: 'line', smooth: false }),
      ],
    }), RECT, 1);

    expect(smooth.beziers).toBeGreaterThan(0);
    expect(straight.beziers).toBe(0);
  });

  it('scales a combo line dash preset by its authored stroke width', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Bars', values: [10, 20, 30], seriesType: 'bar' }),
        series({
          name: 'Sales', values: [100, 120, 160], seriesType: 'line',
          lineColor: 'ED7D31', lineWidthEmu: 31_750,
          chartexStyle: { lineDash: 'dash' }, showMarker: false,
        }),
      ],
    }), RECT, 1);

    expect(rec.strokes).toContainEqual(expect.objectContaining({
      ss: '#ED7D31', lw: 2.5, dash: [15, 7.5],
    }));
  });

  it('keeps axes solid after a dashed combo series and ends its legend key on a full dash', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Jan', 'Feb', 'Mar'],
      showLegend: true,
      legendPos: 'b',
      catAxisLineColor: '000000',
      catAxisLineWidthEmu: 12_700,
      valAxisLineColor: '000000',
      valAxisLineWidthEmu: 12_700,
      series: [
        series({ name: 'Volume', values: [1_200, 1_500, 1_800], seriesType: 'bar' }),
        series({
          name: 'Sales', values: [11_500, 18_000, 24_000], seriesType: 'line',
          lineColor: 'ED7D31', lineWidthEmu: 31_750,
          chartexStyle: { lineDash: 'dash' }, showMarker: false,
          useSecondaryAxis: true,
        }),
      ],
      secondaryValAxis: {
        min: null, max: null, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'out',
      },
    }), RECT, 1);

    expect(rec.strokes.filter(stroke => stroke.ss === '#000000')
      .every(stroke => stroke.dash.length === 0)).toBe(true);
    expect(rec.strokes.filter(stroke => stroke.ss === '#aaa')
      .every(stroke => stroke.dash.length === 0)).toBe(true);
    const orangeKeys = rec.strokes.filter(stroke =>
      stroke.ss === '#ED7D31' && stroke.points.length === 2);
    expect(orangeKeys.some(stroke =>
      Math.abs(stroke.points[1].x - stroke.points[0].x) === 37.5
    )).toBe(true);
  });

  it('uses longer major ticks at category boundaries and shorter minor ticks at centers', () => {
    const ticks = segRecordingCtx();
    const model = baseModel({
      chartType: 'clusteredBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
      series: [series({ values: [1, 2, 3, 4, 5] })],
      catAxisLineColor: '123456',
      catAxisLineWidthEmu: 12_700,
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'cross',
      catAxisFontItalic: true,
    });
    renderChart(ticks.ctx, model, RECT, 1);
    const categoryTicks = ticks.segs.filter(segment =>
      segment.ss === '#123456'
      && Math.abs(segment.x1 - segment.x0) < 0.001
      && Math.abs(segment.y1 - segment.y0) <= 8);
    expect(categoryTicks).toHaveLength(11);
    const byLength = (length: number) => categoryTicks.filter(segment =>
      Math.abs(Math.abs(segment.y1 - segment.y0) - length) < 0.001);
    const boundaryTicks = byLength(6).sort((left, right) => left.x0 - right.x0);
    const centreTicks = byLength(4).sort((left, right) => left.x0 - right.x0);
    expect(boundaryTicks).toHaveLength(6);
    expect(centreTicks).toHaveLength(5);
    expect(boundaryTicks[0].x0).toBeLessThan(centreTicks[0].x0);
    expect(boundaryTicks.at(-1)!.x0).toBeGreaterThan(centreTicks.at(-1)!.x0);
    for (let index = 0; index < centreTicks.length; index++) {
      expect(centreTicks[index].x0).toBeCloseTo(
        (boundaryTicks[index].x0 + boundaryTicks[index + 1].x0) / 2,
        6,
      );
    }

    const labels = recordingCtx();
    renderChart(labels.ctx, model, RECT, 1);
    expect(labels.texts.find(text => text.text === 'Jan')?.font).toContain('italic');
  });

  it('crosses a column category-axis rule at value zero while low labels remain below', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ values: [-10, 10] })],
      valMin: -10,
      valMax: 10,
      catAxisCrosses: 'autoZero',
      catAxisTickLabelPos: 'low',
      catAxisLineColor: '123456',
      catAxisLineWidthEmu: 12_700,
    }), RECT, 1);
    const rule = rec.segs.find(segment =>
      segment.ss === '#123456'
      && Math.abs(segment.y1 - segment.y0) < 0.001
      && Math.abs(segment.x1 - segment.x0) > 100);
    expect(rule).toBeDefined();
    expect(rule!.y0).toBeGreaterThan(RECT.h * 0.2);
    expect(rule!.y0).toBeLessThan(RECT.h * 0.8);
    expect(rec.texts.filter(text => text.text === 'A' || text.text === 'B')
      .every(text => text.y > rule!.y0)).toBe(true);
  });

  it('pads a zero-anchored automatic secondary axis from its effective span', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1_200, 1_500, 1_800], seriesType: 'bar' }),
        series({
          values: [11_500, 18_000, 24_000], seriesType: 'line',
          useSecondaryAxis: true, showMarker: false,
        }),
      ],
      secondaryValAxis: {
        min: null, max: null, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'out',
      },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '25000')).toBe(true);
    expect(rec.texts.some(text => text.text === '30000')).toBe(true);
  });

  it('bar + area combo paints the area as a filled path behind the columns', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Jan', 'Feb', 'Mar'],
      showLegend: true,
      series: [
        series({ name: 'A', values: [45, 52, 30], color: '4472C4', seriesType: 'bar' }),
        series({ name: 'B', values: [25, 30, 45], color: 'A5A5A5', seriesType: 'bar' }),
        series({
          name: 'Trend', values: [80, 85, 90], color: '70AD47',
          seriesType: 'area', useSecondaryAxis: true,
        }),
      ],
      secondaryValAxis: {
        min: 0,
        max: 120,
        title: 'Also Values',
        hidden: false,
        lineHidden: false,
        majorTickMark: 'out',
      },
    }), RECT, 1);

    expect(rec.filledPaths.some(path => path.fillStyle === '#70AD47')).toBe(true);
    expect(rec.rects.filter(rect => rect.fs === '#4472C4' && rect.h > 10)).toHaveLength(3);
    expect(rec.rects.filter(rect => rect.fs === '#A5A5A5' && rect.h > 10)).toHaveLength(3);
    expect(rec.texts.some(text => text.text === 'Trend')).toBe(true);
    expect(rec.texts.some(text => text.text === '120')).toBe(true);
    expect(rec.texts.some(text => text.text === '160')).toBe(false);
  });

  it('uses the authored alternate fill for negative bars', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Positive', 'Negative'],
      series: [series({
        name: 'Profit',
        color: '4472C4',
        values: [150, -300],
        invertIfNegative: true,
        invertedFill: { fillType: 'solid', color: 'FFFFFF' },
        invertedLineColor: '000000',
        invertedLineWidthEmu: 9_525,
      })],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#4472C4')).toHaveLength(1);
    expect(rec.rects.filter(rect => rect.fs === '#FFFFFF')).toHaveLength(1);
    expect(rec.strokeRects.filter(rect => rect.ss === '#000000' && rect.lw === 0.75))
      .toHaveLength(1);
  });

  it('limits an omitted alternate outline below authored line ownership', () => {
    const render = (legacyChartStyle: number, overrides: Partial<ChartSeries> = {}) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar',
        legacyChartStyle,
        categories: ['Negative'],
        series: [series({
          values: [-3],
          invertIfNegative: true,
          invertedFill: { fillType: 'solid', color: 'FFFFFF' },
          invertedFillAuthored: true,
          invertedLineAuthored: false,
          ...overrides,
        })],
      }), RECT, 1);
      return rec;
    };

    expect(render(2).strokeRects.filter(rect => rect.ss === '#000000' && rect.lw === 0.75))
      .toHaveLength(1);
    expect(render(10).strokeRects.filter(rect => rect.ss === '#000000' && rect.lw === 0.75))
      .toHaveLength(0);
    const directLine = render(2, { lineColor: 'FF0000', lineWidthEmu: 12_700 });
    expect(directLine.strokeRects.filter(rect => rect.ss === '#FF0000' && rect.lw === 1))
      .toHaveLength(1);
    expect(directLine.strokeRects.filter(rect => rect.ss === '#000000')).toHaveLength(0);
    expect(render(2, { lineHidden: true }).strokeRects).toHaveLength(0);
    expect(render(2, {
      invertedLineAuthored: true,
      invertedLineColor: '00AA00',
      invertedLineWidthEmu: 25_400,
    }).strokeRects.filter(rect => rect.ss === '#00AA00' && rect.lw === 2)).toHaveLength(1);
    expect(render(2, {
      invertedLineAuthored: true,
      invertedLineHidden: true,
    }).strokeRects).toHaveLength(0);
  });

  it('keeps the application-generated outline-only negative style separate from authored inversion', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C'],
      series: [series({
        name: 'Value',
        color: '4472C4',
        values: [-24_000, -18_000, -11_500],
        automaticNegativeStyle: true,
      })],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#4472C4')).toHaveLength(0);
    expect(rec.strokeRects.filter(rect => rect.ss === '#000000' && rect.lw === 0.75))
      .toHaveLength(3);
  });
});

describe('classic chart data table (CT_DTable)', () => {
  it('reserves a table band and paints keys, names, values, and authored borders', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['45658', '45689'],
      catAxisFormatCode: 'mmm\\-yy',
      series: [
        series({ name: 'Sales', values: [120, 150], color: '4F81BD', seriesType: 'bar' }),
        series({
          name: 'Growth', values: [0.02, 0.05], color: 'C0504D', lineColor: 'C0504D',
          seriesType: 'line', showMarker: true,
        }),
      ],
      dataTable: {
        showHorizontalBorder: true,
        showVerticalBorder: true,
        showOutline: true,
        showKeys: true,
        fontSizeHpt: 1000,
        lineColor: '445566',
        lineWidthEmu: 12700,
        lineDash: 'dash',
        fillColor: 'FFF2CC',
      },
    }), RECT, 1);

    const text = rec.texts.map(call => call.text);
    expect(text).toContain('Sales');
    expect(text).toContain('Growth');
    expect(text).toContain('120');
    expect(text).toContain('0.05');
    expect(text).toContain('Jan-25');
    expect(text).toContain('Feb-25');
    expect(rec.strokeRects.some(rect => rect.ss === '#445566')).toBe(true);
    expect(rec.strokeRects.find(rect => rect.ss === '#445566')?.dash.length).toBeGreaterThan(0);
    const textBackgrounds = rec.rects.filter(rect => rect.fs === '#FFF2CC');
    expect(textBackgrounds.length).toBeGreaterThan(0);
    const sales = rec.texts.find(call => call.text === 'Sales') as NonNullable<typeof rec.texts[number]>;
    expect(textBackgrounds.some(rect =>
      rect.x <= sales.x && rect.x + rect.w >= sales.x
      && rect.y <= sales.y && rect.y + rect.h >= sales.y,
    )).toBe(false);
    expect(rec.arcs.length).toBeGreaterThan(0); // line-series key marker
  });

  it('honors each authored border switch and an explicit noFill line independently', () => {
    const render = (over: Partial<NonNullable<ChartModel['dataTable']>>) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['Q1', 'Q2'],
        series: [series({ name: 'North', values: [10, 15], seriesType: 'line' })],
        dataTable: {
          showHorizontalBorder: false,
          showVerticalBorder: false,
          showOutline: false,
          showKeys: false,
          lineColor: '123ABC',
          ...over,
        },
      }), RECT, 1);
      return {
        lineStrokes: rec.paintEvents.filter(
          event => event.kind === 'stroke' && event.strokeStyle === '#123ABC',
        ).length,
        outlines: rec.strokeRects.filter(rect => rect.ss === '#123ABC').length,
      };
    };

    const baseline = render({});
    expect(baseline.outlines).toBe(0);
    expect(render({ showHorizontalBorder: true }).lineStrokes).toBeGreaterThan(baseline.lineStrokes);
    expect(render({ showVerticalBorder: true }).lineStrokes).toBeGreaterThan(baseline.lineStrokes);
    expect(render({ showOutline: true }).outlines).toBe(1);
    expect(render({
      showHorizontalBorder: true,
      showVerticalBorder: true,
      showOutline: true,
      lineHidden: true,
    })).toEqual({ lineStrokes: 0, outlines: 0 });
  });

  it('applies the text-background semantic independently of chart layout details', () => {
    const renderFillRects = (over: Partial<ChartModel>) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['Q1', 'Q2'],
        series: [series({ name: 'North', values: [10, 15], seriesType: 'line' })],
        dataTable: {
          showHorizontalBorder: true,
          showVerticalBorder: true,
          showOutline: true,
          showKeys: false,
          fillColor: 'FFF2CC',
        },
        ...over,
      }), RECT, 1);
      return rec.rects.filter(rect => rect.fs === '#FFF2CC');
    };

    expect(renderFillRects({}).length).toBeGreaterThan(0);
    expect(renderFillRects({
      plotGroups: [{
        kind: 'line',
        seriesStart: 0,
        seriesCount: 1,
        categoryAxis: 'primary',
        valueAxis: 'primary',
        seriesAxis: 'none',
      }],
    }).length).toBeGreaterThan(0);
    expect(renderFillRects({
      series: [
        series({ name: 'North', values: [10, 15], seriesType: 'line' }),
        series({ name: 'South', values: [8, 12], seriesType: 'bar' }),
      ],
      plotGroups: [
        {
          kind: 'line',
          seriesStart: 0,
          seriesCount: 1,
          categoryAxis: 'primary',
          valueAxis: 'primary',
          seriesAxis: 'none',
        },
        {
          kind: 'bar',
          seriesStart: 1,
          seriesCount: 1,
          categoryAxis: 'primary',
          valueAxis: 'primary',
          seriesAxis: 'none',
          barDirection: 'col',
        },
      ],
    }).length).toBeGreaterThan(0);
    expect(renderFillRects({
      series: [
        series({ name: 'North', values: [10, 15], seriesType: 'line' }),
        series({ name: 'South', values: [8, 12], seriesType: 'bar' }),
      ],
    }).length).toBeGreaterThan(0);
    expect(renderFillRects({ chartType: 'clusteredBarH' }).length).toBeGreaterThan(0);
    expect(renderFillRects({ chartType: 'stock' }).length).toBeGreaterThan(0);
    expect(renderFillRects({ plotAreaManualLayout: { x: 0.2, y: 0.2 } }).length)
      .toBeGreaterThan(0);
    expect(renderFillRects({ categories: [
      ['First quarter with wrapped text repeated', 'until it exceeds the category column'].join(' '),
      'Q2',
    ] }).length)
      .toBeGreaterThan(0);
    expect(renderFillRects({
      series: [series({ name: 'North', values: [10, null], seriesType: 'line' })],
    }).length).toBeGreaterThan(0);
  });

  it('uses the linked dataTable line role only for omitted grid properties', () => {
    const render = (lineColor?: string, roleHidden = false) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'line',
        categories: ['Q1'],
        series: [series({ name: 'North', values: [10], seriesType: 'line' })],
        dataTable: {
          showHorizontalBorder: false,
          showVerticalBorder: false,
          showOutline: true,
          showKeys: false,
          lineColor,
        },
        chartStyleRoles: {
          dataTable: {
            lineColors: ['AABBCC'],
            lineWidthEmu: 19_050,
            lineDash: 'dash',
            lineHidden: roleHidden,
          },
        },
      }), RECT, 1);
      return rec.strokeRects;
    };

    const linked = render();
    expect(linked).toContainEqual(expect.objectContaining({ ss: '#AABBCC', lw: 1.5 }));
    expect(linked.find(rect => rect.ss === '#AABBCC')?.dash.length).toBeGreaterThan(0);

    const direct = render('112233');
    expect(direct.some(rect => rect.ss === '#112233')).toBe(true);
    expect(direct.some(rect => rect.ss === '#AABBCC')).toBe(false);
    expect(render(undefined, true)).toHaveLength(0);
    expect(render('112233', true).some(rect => rect.ss === '#112233')).toBe(true);
  });

  it('localizes a built-in short-date category source independently of the date-axis code', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['45658', '45688'],
      catAxisFormatCode: 'mmm\\-yy',
      series: [series({
        name: 'Sales', values: [120, 150], seriesType: 'bar',
        catFormatCode: 'm/d/yy', catFormatBuiltinId: 14,
      })],
      dataTable: {
        showHorizontalBorder: true, showVerticalBorder: true,
        showOutline: true, showKeys: false,
      },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text.includes('2025'))).toBe(true);
    expect(rec.texts.some(text => text.text === 'Jan-25')).toBe(false);
  });

  it.each(['line', 'area', 'stock'] as const)(
    'uses the shared measured data-table path for %s charts',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['Q1', 'Q2'],
        series: [
          series({ name: 'North', values: [10, 15], seriesType: chartType }),
          series({ name: 'South', values: [18, 13], seriesType: chartType }),
          ...(chartType === 'stock'
            ? [series({ name: 'Close', values: [14, 17], seriesType: 'stock' })]
            : []),
        ],
        dataTable: {
          showHorizontalBorder: true,
          showVerticalBorder: true,
          showOutline: true,
          showKeys: true,
          lineColor: 'C00000',
          lineDash: 'dash',
        },
      }), RECT, 1);

      const text = rec.texts.map(call => call.text);
      expect(text).toContain('North');
      expect(text).toContain('South');
      expect(text).toContain('Q1');
      expect(text).toContain('10');
      expect(rec.strokeRects.some(rect => rect.ss === '#C00000')).toBe(true);
    },
  );

  it('suppresses the duplicate category-axis labels when the table owns the category header', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Q1', 'Q2'],
      series: [series({ name: 'North', values: [10, 15], seriesType: 'line' })],
      dataTable: {
        showHorizontalBorder: true,
        showVerticalBorder: true,
        showOutline: true,
        showKeys: true,
      },
    }), RECT, 1);

    expect(rec.texts.filter(call => call.text === 'Q1')).toHaveLength(1);
    expect(rec.texts.filter(call => call.text === 'Q2')).toHaveLength(1);
  });

  it('keeps horizontal-bar category labels and uses the Office table row order', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Q1', 'Q2'],
      series: [
        series({ name: 'North', values: [10, 15], seriesType: 'bar' }),
        series({ name: 'South', values: [18, 13], seriesType: 'bar' }),
      ],
      dataTable: {
        showHorizontalBorder: true,
        showVerticalBorder: true,
        showOutline: true,
        showKeys: true,
      },
    }), RECT, 1);

    expect(rec.texts.filter(call => call.text === 'Q1')).toHaveLength(2);
    const names = rec.texts
      .map(call => call.text)
      .filter(text => text === 'North' || text === 'South');
    expect(names.slice(0, 2)).toEqual(['South', 'North']);
  });

  it('attaches the shared table to an authored inner plot with a secondary-axis series', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Q1', 'Q2'],
      series: [
        series({ name: 'Primary', values: [10, 15], seriesType: 'line' }),
        series({
          name: 'Secondary', values: [100, 150], seriesType: 'line', useSecondaryAxis: true,
        }),
      ],
      secondaryValAxis: {
        min: 0,
        max: 200,
        title: null,
        hidden: false,
        lineHidden: false,
        majorTickMark: 'out',
      },
      plotAreaBg: 'ABCDEF',
      plotAreaManualLayout: {
        layoutTarget: 'inner',
        xMode: 'factor',
        yMode: 'factor',
        wMode: 'factor',
        hMode: 'factor',
        x: 0.18,
        y: 0.12,
        w: 0.7,
        h: 0.48,
      },
      dataTable: {
        showHorizontalBorder: true,
        showVerticalBorder: true,
        showOutline: true,
        showKeys: true,
      },
    }), RECT, 1);

    const plot = rec.rects.find(rect => rect.fs === '#ABCDEF');
    const header = rec.texts.find(call => call.text === 'Q1');
    expect(plot).toBeDefined();
    expect(header?.y).toBeGreaterThanOrEqual((plot?.y ?? 0) + (plot?.h ?? 0));
    expect(rec.texts.some(call => call.text === 'Primary')).toBe(true);
    expect(rec.texts.some(call => call.text === 'Secondary')).toBe(true);
  });

  it('does not invent a data table for scatter because Office ignores CT_DTable on scatter plots', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: ['1', '2'],
      series: [series({ name: 'North', values: [10, 15], seriesType: 'scatter' })],
      dataTable: {
        showHorizontalBorder: true,
        showVerticalBorder: true,
        showOutline: true,
        showKeys: true,
      },
    }), RECT, 1);

    expect(rec.texts.some(call => call.text === 'North')).toBe(false);
  });
});

describe('classic multi-level category labels', () => {
  it('paints sparse outer labels centered across their category spans', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Male', 'Female', 'Male', 'Female'],
      categoryLevels: [
        ['Male', 'Female', 'Male', 'Female'],
        ['Smoker', '', 'Non-Smoker', ''],
      ],
      series: [series({ name: 'Prevalence', values: [25, 22, 15, 18] })],
    }), RECT, 1);

    const male = rec.texts.find(text => text.text === 'Male');
    const smoker = rec.texts.find(text => text.text === 'Smoker');
    const nonSmoker = rec.texts.find(text => text.text === 'Non-Smoker');
    expect(smoker?.y).toBeGreaterThan(male?.y as number);
    expect(smoker?.x).toBeLessThan(nonSmoker?.x as number);
    expect(rec.paintEvents.filter(event => event.kind === 'stroke').length).toBeGreaterThan(4);
  });

  it('extends first-level separators through the first label band', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Male', 'Female', 'Male', 'Female'],
      categoryLevels: [
        ['Male', 'Female', 'Male', 'Female'],
        ['Smoker', '', 'Non-Smoker', ''],
      ],
      catAxisFontSizeHpt: 1000,
      catAxisLineColor: '000000',
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'cross',
      series: [series({ name: 'Prevalence', values: [25, 22, 15, 18] })],
    }), RECT, 1);

    const separatorSegments = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) > 10);
    const separatorXs = new Set(
      separatorSegments.map(segment => Math.round(segment.x0 * 100) / 100),
    );
    expect(separatorXs.size).toBe(5);
    expect(separatorSegments).toHaveLength(5);
  });

  it('paints each multi-level boundary once with one continuous axis stroke', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Male', 'Female', 'Male', 'Female'],
      categoryLevels: [
        ['Male', 'Female', 'Male', 'Female'],
        ['Smoker', '', 'Non-Smoker', ''],
      ],
      catAxisFontSizeHpt: 1000,
      catAxisLineColor: '000000',
      catAxisLineWidthEmu: 12_700,
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'none',
      valAxisLineHidden: true,
      valAxisMajorGridlines: false,
      series: [series({ name: 'Prevalence', values: [25, 22, 15, 18] })],
    }), RECT, 1);

    const categoryRule = rec.segs.find(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100);
    expect(categoryRule).toBeDefined();
    const axisY = categoryRule!.y0;
    const boundaries = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.min(segment.y0, segment.y1) < axisY
      && Math.max(segment.y0, segment.y1) > axisY + 10);
    expect(boundaries).toHaveLength(5);
    expect(boundaries.every(segment => segment.lw === 1)).toBe(true);
    const allAxisCrossingVertical = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.min(segment.y0, segment.y1) < axisY
      && Math.max(segment.y0, segment.y1) > axisY);
    expect(allAxisCrossingVertical).toHaveLength(5);
  });

  it('keeps major ticks on the actual axis when labels and brackets are low', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      categoryLevels: [
        ['A', 'B', 'C', 'D'],
        ['Left', '', 'Right', ''],
      ],
      catAxisTickLabelPos: 'low',
      catAxisLineColor: '000000',
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'none',
      valAxisLineHidden: true,
      valAxisMajorGridlines: false,
      series: [series({ name: 'Mixed', values: [-10, 15, -5, 20] })],
    }), RECT, 1);

    const categoryRule = rec.segs.find(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100);
    expect(categoryRule).toBeDefined();
    const axisY = categoryRule!.y0;
    const actualAxisTicks = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.min(segment.y0, segment.y1) < axisY
      && Math.max(segment.y0, segment.y1) > axisY
      && Math.abs(segment.y1 - segment.y0) < 20);
    expect(actualAxisTicks).toHaveLength(5);
  });

  it('extends only authored major boundaries into the plot when ticks are skipped', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B', 'C', 'D'],
      categoryLevels: [
        ['A', 'B', 'C', 'D'],
        ['Left', '', 'Right', ''],
      ],
      catAxisLineColor: '000000',
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'none',
      catAxisTickMarkSkip: 2,
      valAxisLineHidden: true,
      valAxisMajorGridlines: false,
      series: [series({ name: 'Positive', values: [10, 15, 5, 20] })],
    }), RECT, 1);

    const categoryRule = rec.segs.find(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100);
    expect(categoryRule).toBeDefined();
    const axisY = categoryRule!.y0;
    const plotwardBoundaries = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.min(segment.y0, segment.y1) < axisY
      && Math.max(segment.y0, segment.y1) > axisY + 10);
    expect(plotwardBoundaries).toHaveLength(3);
  });

  it('does not revive hidden category ticks above multi-level brackets', () => {
    const bracketGeometry = (majorTickMark: 'cross' | 'none') => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBar',
        categories: ['A', 'B', 'C', 'D'],
        categoryLevels: [
          ['A', 'B', 'C', 'D'],
          ['Left', '', 'Right', ''],
        ],
        catAxisLineHidden: true,
        catAxisMajorTickMark: majorTickMark,
        catAxisMinorTickMark: 'none',
        valAxisLineHidden: true,
        valAxisMajorGridlines: false,
        series: [series({ name: 'Positive', values: [10, 15, 5, 20] })],
      }), RECT, 1);
      return rec.segs
        .filter(segment =>
          Math.abs(segment.x1 - segment.x0) < 0.01
          && Math.abs(segment.y1 - segment.y0) > 10)
        .map(segment => [segment.x0, segment.y0, segment.x1, segment.y1]);
    };

    const withoutTicks = bracketGeometry('none');
    expect(withoutTicks).toHaveLength(5);
    expect(bracketGeometry('cross')).toEqual(withoutTicks);
  });

  it('keeps horizontal-bar major ticks when category levels are present', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH',
      categories: ['Male', 'Female', 'Male', 'Female'],
      categoryLevels: [
        ['Male', 'Female', 'Male', 'Female'],
        ['Smoker', '', 'Non-Smoker', ''],
      ],
      catAxisLineColor: '000000',
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'none',
      valAxisHidden: true,
      valAxisMajorGridlines: false,
      series: [series({ name: 'Prevalence', values: [25, 22, 15, 18] })],
    }), RECT, 1);

    const majorTicks = rec.segs.filter(segment =>
      segment.ss === '#000000'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 2
      && Math.abs(segment.x1 - segment.x0) < 20);
    expect(majorTicks).toHaveLength(5);
  });

  it('honors noMultiLvlLbl by suppressing outer labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Male', 'Female'],
      categoryLevels: [['Male', 'Female'], ['Group', '']],
      catAxisNoMultiLevelLabels: true,
      series: [series({ values: [1, 2] })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === 'Group')).toBe(false);
  });
});

describe('CH9 — dispBlanksAs controls null-cell handling (§21.2.2.42)', () => {
  // A series with a hole in the middle: gap breaks the line, zero pins the
  // point to the value-axis zero, span bridges the neighbours with a straight
  // line (the null is skipped, the two sides connect).
  function holeModel(chartType: 'line', dispBlanksAs?: string): ChartModel {
    return baseModel({
      chartType,
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, null, 20] })],
      ...(dispBlanksAs ? { dispBlanksAs } : {}),
    });
  }

  /** The single plotted-line segment (the polyline the series stroked). Chrome
   *  (gridlines/axis) is flat in one axis; the data line varies in both. */
  function dataLine(segs: Array<Array<{ x: number; y: number }>>): Array<{ x: number; y: number }> {
    const data = segs.filter(s => {
      if (s.length < 2) return false;
      const xs = new Set(s.map(p => Math.round(p.x)));
      return xs.size > 1; // spans horizontally → it's the value polyline
    });
    // The longest such segment is the series line.
    return data.sort((a, b) => b.length - a.length)[0] ?? [];
  }

  it('gap (default when absent): the null breaks the line, nothing plots at the middle category', () => {
    // With a middle hole the line must NOT connect A→C directly. The default
    // (no dispBlanksAs) keeps the historical gap behavior (byte-stable).
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line'), RECT, 1);
    const line = dataLine(rec.segments);
    const midX = RECT.x + RECT.w / 2;
    const nearMid = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    // gap: no vertex at the middle category (the null point is skipped and not
    // bridged, so nothing is plotted near the center x from the connecting run).
    expect(nearMid.length).toBe(0);
  });

  it('zero: the null cell plots at the value-axis zero (a low mid vertex)', () => {
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line', 'zero'), RECT, 1);
    const line = dataLine(rec.segments);
    const midX = RECT.x + RECT.w / 2;
    const midPts = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    // zero: the middle category IS plotted (at value 0), so a vertex exists near
    // the center x — and it sits at the BOTTOM of the plot (largest y).
    expect(midPts.length).toBeGreaterThan(0);
    const maxY = Math.max(...line.map(p => p.y));
    expect(midPts.some(p => Math.abs(p.y - maxY) < 1)).toBe(true);
  });

  it('span: the null is skipped but A and C connect directly (no mid vertex, endpoints high)', () => {
    const rec = pathRecordingCtx();
    renderChart(rec.ctx, holeModel('line', 'span'), RECT, 1);
    const line = dataLine(rec.segments);
    // span: only A and C are vertices, joined by a straight lineTo, so the
    // polyline has exactly the two endpoints and NO mid vertex (unlike zero) —
    // yet unlike gap the run is continuous.
    const midX = RECT.x + RECT.w / 2;
    const midPts = line.filter(p => Math.abs(p.x - midX) < RECT.w * 0.1);
    expect(midPts.length).toBe(0);
    // Both endpoints present and at their real (non-zero) heights — the chord
    // runs high across the plot, not down to the baseline.
    const firstX = RECT.x + RECT.w * (0.5 / 3);
    const lastX = RECT.x + RECT.w * (2.5 / 3);
    expect(line.some(p => Math.abs(p.x - firstX) < RECT.w * 0.12)).toBe(true);
    expect(line.some(p => Math.abs(p.x - lastX) < RECT.w * 0.12)).toBe(true);
  });
});

describe('CH9 — dispBlanksAs="zero" applies to per-point data labels too (§21.2.2.42)', () => {
  // The marker loop (line 1452 in renderer.ts) already draws a marker for a
  // null point in "zero" mode. drawCategoryDataLabels must agree: a null cell
  // reads as 0 for BOTH the marker and its label, so "0" is drawn at the null
  // category — matching the spec's "treat the blank cell as zero" semantics
  // (a zero value gets a value label like any other plotted point).
  function labelHoleModel(dispBlanksAs?: string): ChartModel {
    return baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [series({
        name: 'S',
        values: [10, null, 20],
        seriesDataLabels: { showVal: true, showSerName: false, showCatName: false, showPercent: false },
      })],
      ...(dispBlanksAs ? { dispBlanksAs } : {}),
    });
  }

  /** Data labels only — excludes the value-axis tick column (fixed left x) and
   *  the category-axis row (fixed bottom y), which also emit plain numeric /
   *  "A"/"B"/"C" text via fillText. */
  function dataLabelTexts(texts: TextCall[]): string[] {
    const axisTickX = Math.min(...texts.map(t => t.x));
    const catAxisY = Math.max(...texts.map(t => t.y));
    return texts
      .filter(t => Math.abs(t.x - axisTickX) > 1 && Math.abs(t.y - catAxisY) > 1)
      .map(t => t.text);
  }

  it('zero: the null category gets a "0" label alongside 10 and 20', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel('zero'), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts).toContain('0');
  });

  it('gap (default when absent): the null category gets no label at all', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel(), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts.some(t => t === '0')).toBe(false);
  });

  it('span: the null category is skipped (no label), same as gap', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, labelHoleModel('span'), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts.some(t => t === '0')).toBe(false);
  });

  it('a stacked line always labels a null cell at 0, regardless of dispBlanksAs (a stacked sum already reads null as 0)', () => {
    // Mirrors the marker loop's own gate (renderer.ts ~line 1453): stacked
    // series never skip a null point, independent of dispBlanksAs — a null
    // contributes 0 to the running stack sum either way. No dispBlanksAs set
    // (defaults to "gap" for an unstacked series) must NOT suppress the label
    // here, since this series is stacked.
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedLine',
      categories: ['A', 'B', 'C'],
      series: [series({
        name: 'S',
        values: [10, null, 20],
        seriesDataLabels: { showVal: true, showSerName: false, showCatName: false, showPercent: false },
      })],
    }), RECT, 1);
    const labelTexts = dataLabelTexts(rec.texts);
    expect(labelTexts).toContain('10');
    expect(labelTexts).toContain('20');
    expect(labelTexts).toContain('0');
  });
});

// ─── CH8 — pie / doughnut geometry ───────────────────────────────────────────

interface RingArc { x: number; y: number; r: number; a0: number; a1: number; ccw: boolean }
interface FontText { text: string; font: string; fill: string; x: number; y: number }

interface RingRecorded {
  ctx: CanvasRenderingContext2D;
  arcs: RingArc[];
  fills: string[];
  fontTexts: FontText[];
  rotates: number[];
  strokes: Array<{ strokeStyle: string; lineWidth: number }>;
}

/** Recording context that also captures arc() (radius + angles) and, for each
 *  fillText, the active font + fillStyle. Used by the pie/doughnut + font tests
 *  which assert on ring radii, slice start angle, explosion offsets, and the
 *  resolved `ctx.font` family. */
function ringRecordingCtx(): RingRecorded {
  const arcs: RingArc[] = [];
  const fills: string[] = [];
  const fontTexts: FontText[] = [];
  const rotates: number[] = [];
  const strokes: RingRecorded['strokes'] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    lineCap: 'butt', lineJoin: 'miter',
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'arc':
          return (x: number, y: number, r: number, a0: number, a1: number, ccw = false) =>
            arcs.push({ x, y, r, a0, a1, ccw });
        case 'fill':
          return () => fills.push(String(state.fillStyle));
        case 'fillText':
          return (text: string, x: number, y: number) =>
            fontTexts.push({ text, font: String(state.font), fill: String(state.fillStyle), x, y });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'stroke':
          return () => strokes.push({
            strokeStyle: String(state.strokeStyle),
            lineWidth: Number(state.lineWidth),
          });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'moveTo': case 'lineTo': case 'bezierCurveTo':
        case 'quadraticCurveTo': case 'rect': case 'fillRect': case 'strokeRect':
        case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': return () => undefined;
        case 'rotate': return (angle: number) => rotates.push(angle);
        case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default:
          return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    arcs,
    fills,
    fontTexts,
    rotates,
    strokes,
  };
}

/** Outer/inner ring radii for a pie/doughnut: the outer radius is the largest
 *  arc radius; the inner radius is the smallest DISTINCT smaller radius (0 for a
 *  solid pie whose wedges are a single radius). */
function ringRadii(arcs: RingArc[]): { outer: number; inner: number } {
  const rs = [...new Set(arcs.map(a => Math.round(a.r * 100) / 100))].sort((a, b) => b - a);
  return { outer: rs[0] ?? 0, inner: rs.length > 1 ? rs[rs.length - 1] : 0 };
}

describe('CH8 — pie / doughnut geometry', () => {
  const pieModel = (over: Partial<ChartModel>): ChartModel =>
    baseModel({
      chartType: 'pie',
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [30, 45, 25] })],
      ...over,
    });

  it('a plain pie draws solid wedges (inner radius 0)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({}), RECT, 1);
    const { outer, inner } = ringRadii(rec.arcs);
    expect(outer).toBeGreaterThan(0);
    expect(inner).toBe(0);
  });

  it('doughnut holeSize sets the inner radius fraction of the outer radius', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({ chartType: 'doughnut', holeSize: 60 }), RECT, 1);
    const { outer, inner } = ringRadii(rec.arcs);
    expect(inner).toBeGreaterThan(0);
    // holeSize 60 → inner ≈ 0.60 × outer.
    expect(inner / outer).toBeCloseTo(0.6, 2);
  });

  it('a smaller holeSize yields a smaller hole', () => {
    const big = ringRecordingCtx();
    const small = ringRecordingCtx();
    renderChart(big.ctx, pieModel({ chartType: 'doughnut', holeSize: 80 }), RECT, 1);
    renderChart(small.ctx, pieModel({ chartType: 'doughnut', holeSize: 20 }), RECT, 1);
    expect(ringRadii(big.arcs).inner).toBeGreaterThan(ringRadii(small.arcs).inner);
  });

  it('firstSliceAngle rotates the first slice start clockwise from 12 o\'clock', () => {
    const base = ringRecordingCtx();
    const rot = ringRecordingCtx();
    renderChart(base.ctx, pieModel({}), RECT, 1);
    renderChart(rot.ctx, pieModel({ firstSliceAngle: 90 }), RECT, 1);
    // The first wedge's start angle. Default 0 → -90° (canvas up = -π/2).
    const startBase = base.arcs[0].a0;
    const startRot = rot.arcs[0].a0;
    expect(startBase).toBeCloseTo(-Math.PI / 2, 4);
    // +90° → -π/2 + π/2 = 0 (3 o'clock).
    expect(startRot).toBeCloseTo(0, 4);
  });

  it('a transparent hole is NOT overpainted with an opaque fill (doughnut)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({ chartType: 'doughnut', holeSize: 50 }), RECT, 1);
    // Pre-CH8 drew a full 0..2π white circle to mask the wedge centers. The
    // annular geometry removes it: no arc should be a full circle at the inner
    // radius drawn with a white fill immediately after.
    const fullCircles = rec.arcs.filter(a => Math.abs((a.a1 - a.a0) - Math.PI * 2) < 1e-6);
    expect(fullCircles.length).toBe(0);
  });

  it('explosion offsets the slice center outward (arc center moves)', () => {
    const base = ringRecordingCtx();
    renderChart(base.ctx, pieModel({
      series: [series({ name: 'S', values: [30, 45, 25] })],
    }), RECT, 1);
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        values: [30, 45, 25],
        dataPointOverrides: [{ idx: 1, explosion: 40 }],
      })],
    }), RECT, 1);
    // Every wedge shares the pie center EXCEPT the exploded one, whose arc
    // center is displaced. Collect the distinct arc centers.
    const centers = new Set(rec.arcs.map(a => `${Math.round(a.x)},${Math.round(a.y)}`));
    expect(centers.size).toBeGreaterThan(1);
    // The non-exploded pie's shared center — every arc (all 3 slices) is drawn
    // around this single point.
    const trueCenter = base.arcs[0];
    expect(base.arcs.every(a => a.x === trueCenter.x && a.y === trueCenter.y)).toBe(true);
    // Slice 0 and slice 2 (not exploded) still share the true center in the
    // exploded render — only slice 1 moves.
    const outerR = Math.max(...rec.arcs.map(a => a.r));
    const slice0Arcs = rec.arcs.filter(a => a.a0 === base.arcs[0].a0 && a.a1 === base.arcs[0].a1);
    expect(slice0Arcs.length).toBeGreaterThan(0);
    for (const a of slice0Arcs) {
      expect(a.x).toBeCloseTo(trueCenter.x, 6);
      expect(a.y).toBeCloseTo(trueCenter.y, 6);
    }
    // Slice 1 (idx 1, explosion 40): §21.2.2.61 explosion, interpreted (de facto,
    // see ChartDataPointOverride.explosion) as a percentage of the outer radius
    // the slice is displaced outward along its own mid-angle.
    // Values [30, 45, 25] over 2π starting at -π/2 (12 o'clock, clockwise) put
    // slice 1's span at [-π/2 + 0.6π, -π/2 + 1.5π]; its mid-angle is -π/2 + 1.05π.
    const total = 100;
    const startAngle = -Math.PI / 2;
    const slice0Frac = 30 / total;
    const slice1Frac = 45 / total;
    const midAngle = startAngle + slice0Frac * 2 * Math.PI + (slice1Frac * 2 * Math.PI) / 2;
    const expectedOffset = 0.4 * outerR;
    const expectedX = trueCenter.x + Math.cos(midAngle) * expectedOffset;
    const expectedY = trueCenter.y + Math.sin(midAngle) * expectedOffset;
    const slice1Arc = rec.arcs.find(a => Math.abs(a.x - trueCenter.x) > 1 || Math.abs(a.y - trueCenter.y) > 1);
    expect(slice1Arc).toBeDefined();
    expect(slice1Arc?.x).toBeCloseTo(expectedX, 4);
    expect(slice1Arc?.y).toBeCloseTo(expectedY, 4);
    // Displacement magnitude is exactly 40% of the outer radius.
    const dist = Math.hypot((slice1Arc?.x ?? 0) - trueCenter.x, (slice1Arc?.y ?? 0) - trueCenter.y);
    expect(dist).toBeCloseTo(expectedOffset, 4);
  });

  it('uses series explosion as the default and lets a point override it', () => {
    const base = ringRecordingCtx();
    renderChart(base.ctx, pieModel({
      series: [series({ name: 'S', values: [30, 45, 25] })],
    }), RECT, 1);
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S', values: [30, 45, 25], explosion: 40,
        dataPointOverrides: [{ idx: 1, explosion: 0 }],
      })],
    }), RECT, 1);

    const baseCenter = base.arcs[0];
    const slice = (index: number) => rec.arcs.find(arc =>
      arc.a0 === base.arcs[index].a0 && arc.a1 === base.arcs[index].a1
    );
    expect(slice(0)).toBeDefined();
    expect(slice(1)).toBeDefined();
    expect(Math.hypot(
      (slice(0)?.x ?? 0) - baseCenter.x,
      (slice(0)?.y ?? 0) - baseCenter.y,
    )).toBeGreaterThan(1);
    expect(slice(1)?.x).toBeCloseTo(baseCenter.x, 6);
    expect(slice(1)?.y).toBeCloseTo(baseCenter.y, 6);
  });

  it('a multi-series doughnut draws concentric rings (multiple distinct radii)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'doughnut',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Outer', values: [1, 1] }),
        series({ name: 'Inner', values: [1, 1] }),
      ],
      holeSize: 40,
    }), RECT, 1);
    // Two rings → at least three distinct radii (outer ring outer/inner + inner
    // ring outer/inner, some shared) — assert more than the two a single ring
    // would produce.
    const distinctRadii = new Set(rec.arcs.map(a => Math.round(a.r * 10) / 10));
    expect(distinctRadii.size).toBeGreaterThanOrEqual(3);
    // The single-series doughnut geometry (asserted in the tests above) gives
    // us an independently-derived outer radius for this RECT — reuse it so the
    // band boundaries below aren't just copied from the renderer's own formula.
    const single = ringRecordingCtx();
    renderChart(single.ctx, baseModel({
      chartType: 'doughnut', categories: ['A'], series: [series({ name: 'S', values: [1] })], holeSize: 40,
    }), RECT, 1);
    // Use the RAW (unrounded) outer radius so the derived band boundaries below
    // don't compound `ringRadii`'s rounding into a spurious mismatch.
    const outerR = Math.max(...single.arcs.map(a => a.r));
    const innerR = outerR * 0.4; // holeSize 40 → hole is 40% of the outer radius
    const ringBand = (outerR - innerR) / 2; // band from hole to outer edge, split evenly across 2 rings
    const expectRadiiCloseTo = (arcs: RingArc[], expected: number[]): void => {
      const actual = [...new Set(arcs.map(a => Math.round(a.r * 1000) / 1000))].sort((a, b) => b - a);
      const wanted = [...expected].sort((a, b) => b - a);
      expect(actual.length).toBe(wanted.length);
      actual.forEach((r, i) => expect(r).toBeCloseTo(wanted[i], 2));
    };
    // Each ring draws 2 arcs (outer + inner annulus edge) per category (A, B) →
    // 4 arcs per ring, 8 total. Ring 0 ("Outer" series) is drawn FIRST and
    // occupies the OUTERMOST band.
    expectRadiiCloseTo(rec.arcs.slice(0, 4), [outerR, outerR - ringBand]);
    // Ring 1 ("Inner" series) is drawn SECOND and occupies the band adjacent to
    // the hole; its outer edge meets ring 0's inner edge, its inner edge is the
    // hole radius.
    expectRadiiCloseTo(rec.arcs.slice(4), [outerR - ringBand, innerR]);
  });

  it('rich pie dLbls compose showCatName + showPercent', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        categories: ['Alpha', 'Beta', 'Gamma'],
        values: [30, 45, 25],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: true,
        },
      })],
    }), RECT, 1);
    const texts = rec.fontTexts.map(t => t.text);
    // "Alpha 30%" etc. — category name and percent joined.
    expect(texts.some(t => t.includes('Alpha') && t.includes('30%'))).toBe(true);
  });

  it('rich pie dLbls honor the authored separator and value-cache format', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S',
        categories: ['Alpha', 'Beta'],
        values: [0.43, 0.57],
        valFormatCode: '0%',
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          separator: '\n',
        },
      })],
    }), RECT, 1);
    const texts = rec.fontTexts.map(t => t.text);
    expect(texts).toContain('Alpha');
    expect(texts).toContain('43%');
    expect(texts).not.toContain('Alpha 0.43');
  });

  it('rich pie percent labels honor authored numFmt and display scale', () => {
    const model = pieModel({
      series: [series({
        name: 'S', categories: ['A', 'B'], values: [1, 2],
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: true,
          formatCode: '0.0%', fontSizeHpt: 1200, position: 'ctr', fontFace: 'Pie Face',
        },
      })],
    });
    const normal = ringRecordingCtx();
    renderChart(normal.ctx, model, RECT, 1);
    const scaled = ringRecordingCtx();
    renderChart(scaled.ctx, model, RECT, 2);
    expect(normal.fontTexts.some(text => text.text === '33.3%' && text.font.includes('12px')))
      .toBe(true);
    expect(normal.fontTexts.find(text => text.text === '33.3%')?.font)
      .toContain('"Pie Face"');
    expect(scaled.fontTexts.some(text => text.text === '33.3%' && text.font.includes('24px')))
      .toBe(true);
  });

  it.each(['pie', 'doughnut'] as const)(
    '%s custom rich labels paint inline runs with theme-resolved faces',
    chartType => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType,
        categories: ['A'],
        themeMajorFontLatin: 'Major Theme',
        themeMinorFontLatin: 'Minor Theme',
        series: [series({
          values: [1],
          dataLabelOverrides: [{
            idx: 0,
            text: 'Major Minor',
            position: 'ctr',
            richRuns: [
              { text: 'Major', fontFace: '+mj-lt', color: '112233' },
              { text: ' Minor', fontFace: '+mn-lt', color: '445566' },
            ],
          }],
        })],
      }), RECT, 1);

      expect(rec.texts.find(call => call.text === 'Major'))
        .toMatchObject({ fillStyle: '#112233' });
      expect(rec.texts.find(call => call.text === 'Major')?.font)
        .toContain('"Major Theme"');
      expect(rec.texts.find(call => call.text === ' Minor'))
        .toMatchObject({ fillStyle: '#445566' });
      expect(rec.texts.find(call => call.text === ' Minor')?.font)
        .toContain('"Minor Theme"');
    },
  );

  it('inside pie labels respect tiny-slice capacity', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S', categories: ['TINY_LABEL_LONG', 'Large'], values: [1, 999],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: false,
          position: 'ctr',
        },
      })],
    }), RECT, 1);
    expect(rec.fontTexts.some(text => text.text === 'TINY_LABEL_LONG')).toBe(false);
  });

  it('outside pie labels are bounded, elided, and clipped to chart space', () => {
    const rec = recordingCtx();
    const long = 'Outside '.repeat(300);
    renderChart(rec.ctx, pieModel({
      series: [series({
        name: 'S', categories: [long, 'B'], values: [1, 1],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd',
        },
      })],
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === long)).toBe(false);
    expect(rec.texts.some(text => text.text.includes('…'))).toBe(true);
    expect(rec.clips).toContainEqual(RECT);
  });
});

// ─── CH10 — chart text font faces ────────────────────────────────────────────

describe('CH10 — chart text font faces', () => {
  // No data labels: the only numeric text is then the value-axis ticks, so the
  // `/^[\d.]+$/` filter isolates the value-axis font cleanly (data-label values
  // legitimately use the SEPARATE dataLabelFontFace and would otherwise blur the
  // assertion).
  const barWithLabels = (over: Partial<ChartModel>): ChartModel =>
    baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 20] })],
      valAxisTitle: 'Units',
      ...over,
    });

  it('an explicit value-axis face is used for value-axis tick labels', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ valAxisFontFace: 'Georgia' }), RECT, 1);
    // The value-axis ticks ("0", "5", …) are drawn with the Georgia family.
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Georgia'))).toBe(true);
  });

  it('falls back to the theme body (minor) font when no element face is set', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ themeMinorFontLatin: 'Aptos Narrow' }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(true);
  });

  it('an element face wins over the theme font', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({
      valAxisFontFace: 'Georgia',
      themeMinorFontLatin: 'Aptos Narrow',
    }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.some(f => f.includes('Georgia'))).toBe(true);
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(false);
  });

  it('with no face and no theme, the built-in sans-serif is used (byte-stable)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({}), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    expect(tickFonts.length).toBeGreaterThan(0);
    expect(tickFonts.every(f => f.endsWith('sans-serif') && !f.includes('"'))).toBe(true);
  });

  it('invalid public-model axis and legend font sizes fall back instead of expanding layout', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, barWithLabels({
      valAxisFontSizeHpt: 400_001,
      catAxisFontSizeHpt: Number.POSITIVE_INFINITY,
      legendFontSizeHpt: 99,
      showLegend: true,
    }), RECT, 1);
    const usedSizes = rec.texts.map(text => Number(/([0-9.]+)px/.exec(text.font ?? '')?.[1] ?? 0));
    expect(Math.max(...usedSizes)).toBeLessThan(100);
  });

  it('a `+mn-lt` theme reference face resolves to the theme minor font', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({
      valAxisFontFace: '+mn-lt',
      themeMinorFontLatin: 'Aptos Narrow',
      themeMajorFontLatin: 'Aptos Display',
    }), RECT, 1);
    const tickFonts = rec.fontTexts.filter(t => /^[\d.]+$/.test(t.text)).map(t => t.font);
    // "+mn-lt" must NOT appear literally; it resolves to the minor face.
    expect(tickFonts.some(f => f.includes('Aptos Narrow'))).toBe(true);
    expect(tickFonts.some(f => f.includes('+mn-lt'))).toBe(false);
  });

  it('axis titles use the theme heading (major) font as fallback', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, barWithLabels({ themeMajorFontLatin: 'Aptos Display' }), RECT, 1);
    const titleFont = rec.fontTexts.find(t => t.text === 'Units')?.font;
    expect(titleFont).toBeDefined();
    expect(titleFont).toContain('Aptos Display');
  });
});

// ── CH6 — axis scale model (gridlines / units / logBase / orientation) ───────

interface Seg {
  x0: number; y0: number; x1: number; y1: number; ss: string; lw: number;
  dash: number[]; cap: string; join: string;
}
interface SegRecorded { ctx: CanvasRenderingContext2D; segs: Seg[]; texts: TextCall[] }

function strokedPolylineCtx(): {
  ctx: CanvasRenderingContext2D;
  strokes: Array<{
    points: Array<{ x: number; y: number }>; ss: string; lw: number; dash: number[];
    cap: string; join: string;
  }>;
  texts: TextCall[];
} {
  const strokes: Array<{
    points: Array<{ x: number; y: number }>; ss: string; lw: number; dash: number[];
    cap: string; join: string;
  }> = [];
  const texts: TextCall[] = [];
  let path: Array<{ x: number; y: number }> = [];
  let dash: number[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    lineCap: 'butt', lineJoin: 'miter',
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (text: string) => ({ width: String(text).length * 6 });
        case 'beginPath': return () => { path = []; };
        case 'moveTo': return (x: number, y: number) => { path.push({ x, y }); };
        case 'lineTo': return (x: number, y: number) => { path.push({ x, y }); };
        case 'stroke': return () => {
          if (path.length >= 2) {
            strokes.push({
              points: path.map(point => ({ ...point })),
              ss: String(state.strokeStyle),
              lw: Number(state.lineWidth),
              dash: [...dash],
              cap: String(state.lineCap),
              join: String(state.lineJoin),
            });
          }
        };
        case 'setLineDash': return (value: number[]) => { dash = [...value]; };
        case 'getLineDash': return () => [...dash];
        case 'fillText': return (text: string, x: number, y: number) => {
          texts.push({
            text: String(text), x, y,
            align: String(state.textAlign), baseline: String(state.textBaseline),
          });
        };
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default: return () => undefined;
      }
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  };
  return {
    ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D,
    strokes,
    texts,
  };
}

/** Recording context that captures stroked line SEGMENTS (moveTo→lineTo→stroke)
 *  plus fillText, so gridline presence/orientation can be asserted. */
function segRecordingCtx(): SegRecorded {
  const segs: Seg[] = [];
  const texts: TextCall[] = [];
  let dash: number[] = [];
  let cx = 0, cy = 0, mx = 0, my = 0;
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'moveTo': return (x: number, y: number) => { cx = x; cy = y; mx = x; my = y; };
        case 'lineTo': return (x: number, y: number) => {
          segs.push({
            x0: cx, y0: cy, x1: x, y1: y,
            ss: String(state.strokeStyle), lw: Number(state.lineWidth), dash: [...dash],
            cap: String(state.lineCap), join: String(state.lineJoin),
          });
          cx = x; cy = y;
        };
        case 'fillText': return (text: string, x: number, y: number) =>
          texts.push({
            text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline),
            font: String(state.font), fillStyle: String(state.fillStyle),
          });
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'closePath': return () => { cx = mx; cy = my; };
        case 'save': case 'restore': case 'beginPath': case 'fill': case 'stroke':
        case 'arc': case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect':
        case 'fillRect': case 'strokeRect': case 'clearRect': case 'strokeText':
          return () => undefined;
        case 'setLineDash': return (value?: number[]) => {
          dash = Array.isArray(value) ? [...value] : [];
        };
        case 'getLineDash': return () => [...dash];
        case 'translate': case 'rotate': case 'scale': case 'clip':
        case 'setTransform': case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segs, texts };
}

/** Value-axis MAJOR/MINOR gridlines: near-flat segments spanning the plot width
 *  drawn in the gridline colors (`#e0e0e0` faint or `#aaa` zero line). The
 *  category axis bottom rule is also `#aaa` horizontal, so it's excluded by
 *  dropping the single bottom-most horizontal `#aaa` segment (the axis line). */
function horizGridlines(segs: Seg[]): Seg[] {
  const flat = segs.filter(s => Math.abs(s.y0 - s.y1) < 0.5 && Math.abs(s.x1 - s.x0) > 50);
  const grids = flat.filter(s => s.ss === '#e0e0e0' || s.ss === '#aaa');
  // Drop the bottom-most `#aaa` line (the category axis rule) if present.
  const aaa = grids.filter(s => s.ss === '#aaa');
  if (aaa.length === 0) return grids;
  const maxY = Math.max(...aaa.map(s => s.y0));
  let dropped = false;
  return grids.filter(s => {
    if (!dropped && s.ss === '#aaa' && Math.abs(s.y0 - maxY) < 0.5) { dropped = true; return false; }
    return true;
  });
}

/** Category-axis MAJOR gridlines: near-vertical segments spanning the plot
 *  height (x roughly constant, big y span). Filters by the gridline color so
 *  bar edges / data lines aren't counted. */
function vertGridlines(segs: Seg[], color = '#e0e0e0'): Seg[] {
  return segs.filter(s => Math.abs(s.x0 - s.x1) < 0.5 && Math.abs(s.y1 - s.y0) > 50 && s.ss === color);
}

describe('CH6 — axis scale model', () => {
  const lineModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'line',
    categories: ['A', 'B', 'C'],
    series: [series({ name: 'S', values: [10, 20, 30] })],
    ...over,
  });

  it('valAxisMajorGridlines=false suppresses the value gridlines (labels stay)', () => {
    const on = segRecordingCtx();
    renderChart(on.ctx, lineModel({}), RECT, 1);
    const gridsOn = horizGridlines(on.segs).length;
    expect(gridsOn).toBeGreaterThan(0);

    const off = segRecordingCtx();
    renderChart(off.ctx, lineModel({ valAxisMajorGridlines: false }), RECT, 1);
    // No horizontal gridlines spanning the plot when suppressed.
    expect(horizGridlines(off.segs).length).toBe(0);
    // Tick labels still drawn.
    expect(off.texts.some(t => t.text === '10')).toBe(true);
  });

  it('an explicit valAxisGridlineColor strokes the gridlines in that color (§21.2.2.100)', () => {
    // Flat plot-spanning segments in the explicit gridline color.
    const flatOfColor = (segs: Seg[], color: string): Seg[] =>
      segs.filter(s => Math.abs(s.y0 - s.y1) < 0.5 && Math.abs(s.x1 - s.x0) > 50 && s.ss === color);

    // Default (no explicit gridline color) → the faint #e0e0e0 hairline.
    const def = segRecordingCtx();
    renderChart(def.ctx, lineModel({}), RECT, 1);
    expect(flatOfColor(def.segs, '#e0e0e0').length).toBeGreaterThan(0);
    expect(flatOfColor(def.segs, '#8fa878').length).toBe(0);

    // sample-1 slide 5: accent3 (#8FA878) 0.25 pt gridlines. The renderer strokes
    // every major gridline in that color — no faint #e0e0e0 lines remain — and
    // suppresses the #aaa zero-line emphasis (uniform per PowerPoint).
    const styled = segRecordingCtx();
    renderChart(styled.ctx, lineModel({ valAxisGridlineColor: '8fa878', valAxisGridlineWidthEmu: 3175 }), RECT, 1);
    const colored = flatOfColor(styled.segs, '#8fa878');
    expect(colored.length).toBeGreaterThan(0);
    expect(flatOfColor(styled.segs, '#e0e0e0').length).toBe(0);
    // Same gridline COUNT as the default — only the stroke style changed.
    // The default splits its gridlines across #e0e0e0 (non-zero) and a single
    // #aaa zero-line; the explicit color unifies all of them into #8fa878, so
    // the count matches `horizGridlines` (which sums both, dropping the
    // cat-axis rule).
    expect(colored.length).toBe(horizGridlines(def.segs).length);
    // Width floors at 0.5 px (0.25 pt × ptToPx=1 = 0.25 px → floored).
    expect(colored.every(s => s.lw === 0.5)).toBe(true);
  });

  it('uses linked major/minor gridline roles behind direct axis formatting', () => {
    const major = segRecordingCtx();
    renderChart(major.ctx, lineModel({
      valAxisMajorGridlines: true,
      chartStyleRoles: {
        gridlineMajor: { lineColors: ['AABBCC'], lineWidthEmu: 19_050 },
      },
    }), RECT, 1);
    const majorLines = major.segs.filter(segment =>
      segment.ss === '#AABBCC' && Math.abs(segment.x1 - segment.x0) > 50
    );
    expect(majorLines.length).toBeGreaterThan(0);
    expect(majorLines.every(segment => segment.lw === 1.5)).toBe(true);

    const minor = segRecordingCtx();
    renderChart(minor.ctx, lineModel({
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: true,
      valAxisMinorUnit: 2,
      chartStyleRoles: {
        gridlineMinor: { lineColors: ['CCBBAA'], lineWidthEmu: 12_700 },
      },
    }), RECT, 1);
    expect(minor.segs.some(segment =>
      segment.ss === '#CCBBAA' && Math.abs(segment.x1 - segment.x0) > 50 && segment.lw === 1
    )).toBe(true);
  });

  it('keeps direct gridline paint and noFill ahead of linked roles', () => {
    const direct = segRecordingCtx();
    renderChart(direct.ctx, lineModel({
      valAxisMajorGridlines: true,
      valAxisGridlineColor: '112233',
      chartStyleRoles: { gridlineMajor: { lineColors: ['AABBCC'], lineHidden: true } },
    }), RECT, 1);
    expect(direct.segs.some(segment =>
      segment.ss === '#112233' && Math.abs(segment.x1 - segment.x0) > 50
    )).toBe(true);
    expect(direct.segs.some(segment => segment.ss === '#AABBCC')).toBe(false);

    const hidden = segRecordingCtx();
    renderChart(hidden.ctx, lineModel({
      valAxisMajorGridlines: true,
      chartStyleRoles: { gridlineMajor: { lineHidden: true } },
    }), RECT, 1);
    expect(horizGridlines(hidden.segs)).toHaveLength(0);
  });

  it('uses linked category/value-axis line paint behind direct axis formatting', () => {
    const linked = segRecordingCtx();
    renderChart(linked.ctx, lineModel({
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        categoryAxis: { lineColors: ['AABBCC'], lineWidthEmu: 19_050, lineDash: 'dash' },
        valueAxis: { lineColors: ['CCBBAA'], lineWidthEmu: 25_400, lineDash: 'dot' },
      },
    }), RECT, 1);
    expect(linked.segs.some(segment => segment.ss === '#AABBCC' && segment.lw === 1.5)).toBe(true);
    expect(linked.segs.some(segment => segment.ss === '#CCBBAA' && segment.lw === 2)).toBe(true);
    expect(linked.segs.some(segment =>
      segment.ss === '#AABBCC' && segment.dash.length > 0
    )).toBe(true);
    expect(linked.segs.some(segment =>
      segment.ss === '#CCBBAA' && segment.dash.length > 0
    )).toBe(true);

    const direct = segRecordingCtx();
    renderChart(direct.ctx, lineModel({
      valAxisMajorGridlines: false,
      catAxisLineColor: '112233',
      catAxisLineDash: 'sysDot',
      chartStyleRoles: {
        categoryAxis: { lineColors: ['AABBCC'], lineDash: 'dash', lineHidden: true },
      },
    }), RECT, 1);
    expect(direct.segs.some(segment =>
      segment.ss === '#112233' && segment.dash.length === 2
        && segment.dash[0] === 1 && segment.dash[1] === 2
    )).toBe(true);
    expect(direct.segs.some(segment => segment.ss === '#AABBCC')).toBe(false);

    const hidden = segRecordingCtx();
    renderChart(hidden.ctx, lineModel({
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        categoryAxis: { lineColors: ['AABBCC'], lineHidden: true },
        valueAxis: { lineColors: ['CCBBAA'], lineHidden: true },
      },
    }), RECT, 1);
    expect(hidden.segs.some(segment => segment.ss === '#AABBCC' || segment.ss === '#CCBBAA'))
      .toBe(false);
    expect(hidden.texts.map(text => text.text)).toEqual(expect.arrayContaining(['A', '10']));
  });

  it('uses linked axis text defaults behind direct tick-label formatting', () => {
    const linked = segRecordingCtx();
    renderChart(linked.ctx, lineModel({
      valAxisMajorGridlines: false,
      chartStyleRoles: {
        categoryAxis: {
          fontSizeHpt: 700, fontBold: true, fontItalic: true,
          fontColor: 'AABBCC', fontFace: 'Linked Category',
        },
        valueAxis: {
          fontSizeHpt: 800, fontBold: true, fontItalic: true,
          fontColor: 'CCBBAA', fontFace: 'Linked Value',
        },
      },
    }), RECT, 1);
    const category = linked.texts.find(text => text.text === 'A');
    const value = linked.texts.find(text => text.text === '10');
    expect(category).toMatchObject({ fillStyle: '#AABBCC' });
    expect(category?.font).toContain('italic bold 7px "Linked Category"');
    expect(value).toMatchObject({ fillStyle: '#CCBBAA' });
    expect(value?.font).toContain('italic bold 8px "Linked Value"');

    const direct = segRecordingCtx();
    renderChart(direct.ctx, lineModel({
      valAxisMajorGridlines: false,
      catAxisFontSizeHpt: 1100,
      catAxisFontBold: false,
      catAxisFontItalic: false,
      catAxisFontColor: '112233',
      catAxisFontFace: 'Direct Category',
      chartStyleRoles: {
        categoryAxis: {
          fontSizeHpt: 700, fontBold: true, fontItalic: true,
          fontColor: 'AABBCC', fontFace: 'Linked Category',
        },
      },
    }), RECT, 1);
    const directCategory = direct.texts.find(text => text.text === 'A');
    expect(directCategory).toMatchObject({ fillStyle: '#112233' });
    expect(directCategory?.font).toContain('11px "Direct Category"');
    expect(directCategory?.font).not.toContain('italic');
    expect(directCategory?.font).not.toContain('bold');
  });

  it('valAxisTickLabelPos="none" hides value tick labels (gridlines stay)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({ valAxisTickLabelPos: 'none' }), RECT, 1);
    // Value labels (numeric) gone; gridlines still present.
    expect(rec.texts.some(t => /^\d+$/.test(t.text))).toBe(false);
    expect(horizGridlines(rec.segs).length).toBeGreaterThan(0);
  });

  it('an explicit valAxisMajorUnit changes the gridline count', () => {
    // Data 10..30 → auto step 5 (0,5,…,35 ≈ 8 lines). majorUnit 10 → coarser.
    const auto = segRecordingCtx();
    renderChart(auto.ctx, lineModel({}), RECT, 1);
    const coarse = segRecordingCtx();
    renderChart(coarse.ctx, lineModel({ valAxisMajorUnit: 10 }), RECT, 1);
    expect(horizGridlines(coarse.segs).length).toBeLessThan(horizGridlines(auto.segs).length);
    // Labels land on 0,10,20,30,… (multiples of 10) only.
    const coarseLabels = coarse.texts.map(t => t.text).filter(t => /^\d+$/.test(t));
    expect(coarseLabels).toContain('10');
    expect(coarseLabels).toContain('20');
    expect(coarseLabels).not.toContain('5');
  });

  it('valAxisOrientation="maxMin" reverses the value axis (bar heights flip)', () => {
    const normal = recordingCtx();
    renderChart(normal.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 30] })],
    }), RECT, 1);
    const reversed = recordingCtx();
    renderChart(reversed.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [series({ name: 'S', values: [10, 30] })],
      valAxisOrientation: 'maxMin',
    }), RECT, 1);
    // Normal: taller value (30) → shorter y (higher up) and greater height.
    // Reversed: the axis flips, so the bar for 30 grows DOWNWARD from the top.
    const [nSmall, nBig] = normal.rects;
    const [rSmall, rBig] = reversed.rects;
    // In the reversed axis the "30" bar's top edge sits at the plot top area
    // and it extends toward the (now-inverted) zero at the bottom-flipped end;
    // its y origin differs from the normal orientation.
    expect(rBig.y).not.toBeCloseTo(nBig.y, 1);
    // §21.2.2.130 orientation="maxMin" is a true mirror of the value axis, not
    // just "a different y": every value's pixel position reflects across the
    // plot's vertical midline. Both bars are zero-anchored (clustered, single
    // series), so — independent of any internal renderer constant — the
    // reversed zero line is the SHARED top edge of both reversed bars, and the
    // normal zero line is the SHARED bottom edge of both normal bars.
    const reversedZeroY = rSmall.y; // = rBig.y — both bars start at the (flipped) zero line
    expect(rBig.y).toBeCloseTo(reversedZeroY, 6);
    const normalZeroY = nSmall.y + nSmall.h; // = nBig.y + nBig.h — both bars end at zero
    expect(nBig.y + nBig.h).toBeCloseTo(normalZeroY, 6);
    // The mirror axis: for any value v, reversedBottom(v) = 2*reversedZeroY +
    // (normalZeroY - reversedZeroY) - normalTop(v). A reversed bar's BOTTOM
    // edge is the mirror image of the corresponding normal bar's TOP edge
    // around the (reversedZeroY, normalZeroY) span.
    const mirror = (yNormalTop: number): number => 2 * reversedZeroY + (normalZeroY - reversedZeroY) - yNormalTop;
    expect(rSmall.y + rSmall.h).toBeCloseTo(mirror(nSmall.y), 4);
    expect(rBig.y + rBig.h).toBeCloseTo(mirror(nBig.y), 4);
    // The smaller value (10) still produces the smaller bar on the reversed
    // axis too — reversal flips direction, not relative magnitude.
    expect(rSmall.h).toBeLessThan(rBig.h);
  });

  it('valAxisLogBase=10 places gridlines on powers of ten', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({
      categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [1, 10, 100] })],
      valAxisLogBase: 10,
    }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    // Decade tick labels 1 / 10 / 100 present (1000 not required for this range).
    expect(labels).toContain('1');
    expect(labels).toContain('10');
    expect(labels).toContain('100');
  });

  it('scatter X/Y axes share logarithmic and reversed numeric-axis mapping', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      catAxisHidden: true,
      valAxisHidden: true,
      catAxisLogBase: 10,
      valAxisLogBase: 10,
      catAxisOrientation: 'maxMin',
      valAxisOrientation: 'maxMin',
      catAxisMin: 1,
      catAxisMax: 100,
      valMin: 1,
      valMax: 100,
      series: [series({
        categories: ['1', '10', '100'],
        values: [1, 10, 100],
        showMarker: true,
        markerSymbol: 'circle',
        lineHidden: true,
      })],
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(3);
    const [low, middle, high] = rec.arcs;
    expect(low.x).toBeGreaterThan(middle.x);
    expect(middle.x).toBeGreaterThan(high.x);
    expect(low.x - middle.x).toBeCloseTo(middle.x - high.x, 6);
    expect(low.y).toBeLessThan(middle.y);
    expect(middle.y).toBeLessThan(high.y);
    expect(middle.y - low.y).toBeCloseTo(high.y - middle.y, 6);
  });

  it('area category orientation reverses points and their authored labels together', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area',
      categories: ['First', 'Middle', 'Last'],
      catAxisOrientation: 'maxMin',
      series: [series({
        values: [10, 20, 30],
        dataLabelOverrides: [
          { idx: 0, text: 'first', position: 'ctr' },
          { idx: 2, text: 'last', position: 'ctr' },
        ],
      })],
    }), RECT, 1);

    const first = rec.texts.find(text => text.text === 'first');
    const last = rec.texts.find(text => text.text === 'last');
    expect(first?.x).toBeGreaterThan(last?.x as number);
  });

  it('secondary value axes share logarithmic ticks and reversed mapping', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [1, 10, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 1,
        max: 100,
        title: null,
        hidden: false,
        lineHidden: false,
        majorTickMark: 'out',
        logBase: 10,
        orientation: 'maxMin',
      },
    }), RECT, 1);

    const rightTicks = rec.texts.filter(text => text.x > RECT.w * 0.75);
    const one = rightTicks.find(text => text.text === '1');
    const ten = rightTicks.find(text => text.text === '10');
    const hundred = rightTicks.find(text => text.text === '100');
    expect(one?.y).toBeLessThan(ten?.y as number);
    expect(ten?.y).toBeLessThan(hundred?.y as number);
  });

  it.each(['clusteredBar', 'line', 'area'] as const)(
    '%s routes secondary major gridlines through the shared under-data layer',
    (chartType) => {
      const secondary = series({
        name: 'Secondary', values: [20, 60, 100], useSecondaryAxis: true,
        ...(chartType === 'clusteredBar' ? { seriesType: 'line' } : {}),
      });
      const model = baseModel({
        chartType,
        categories: ['A', 'B', 'C'],
        series: [series({ name: 'Primary', values: [10, 20, 30] }), secondary],
        valAxisMajorGridlines: false,
        secondaryValAxis: {
          min: 0, max: 100, title: null, hidden: false, lineHidden: false,
          majorTickMark: 'none', majorUnit: 20,
          majorGridlines: true, majorGridlineColor: '654321',
          majorGridlineWidthEmu: 25400,
        },
      });
      const rec = segRecordingCtx();
      renderChart(rec.ctx, model, RECT, 1);

      const grids = rec.segs.filter(segment =>
        segment.ss === '#654321' && Math.abs(segment.x1 - segment.x0) > 50
      );
      expect(grids.length).toBeGreaterThanOrEqual(6);
      expect(grids.every(segment => segment.lw === 2)).toBe(true);

      const order = recordingCtx();
      renderChart(order.ctx, model, RECT, 1);
      const firstGrid = order.paintEvents.findIndex(event =>
        event.kind === 'stroke' && event.strokeStyle === '#654321'
      );
      const firstSeries = order.paintEvents.findIndex(event =>
        event.kind === 'stroke'
        && (event.strokeStyle === '#4472C4' || event.strokeStyle === '#ED7D31')
      );
      expect(firstGrid).toBeGreaterThanOrEqual(0);
      expect(firstSeries).toBeGreaterThan(firstGrid);
    },
  );

  it('applies the linked major-gridline role to an enabled secondary axis', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({
      valAxisMajorGridlines: false,
      series: [
        series({ values: [10, 20, 30] }),
        series({ values: [20, 60, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none', majorUnit: 20, majorGridlines: true,
      },
      chartStyleRoles: {
        gridlineMajor: { lineColors: ['654321'], lineWidthEmu: 25_400 },
      },
    }), RECT, 1);
    const lines = rec.segs.filter(segment =>
      segment.ss === '#654321' && Math.abs(segment.x1 - segment.x0) > 50
    );
    expect(lines.length).toBeGreaterThanOrEqual(6);
    expect(lines.every(segment => segment.lw === 2)).toBe(true);
  });

  it('applies the linked value-axis dash to an unauthored secondary axis', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({
      valAxisMajorGridlines: false,
      valAxisLineColor: '111111',
      valAxisLineDash: 'solid',
      series: [
        series({ values: [10, 20, 30] }),
        series({ values: [20, 60, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none', majorUnit: 20,
      },
      chartStyleRoles: {
        valueAxis: {
          lineColors: ['654321'], lineWidthEmu: 25_400, lineDash: 'dashDot',
          fontSizeHpt: 700, fontItalic: true, fontColor: 'AABBCC', fontFace: 'Secondary Face',
        },
      },
    }), RECT, 1);
    expect(rec.segs.some(segment =>
      segment.ss === '#654321' && segment.lw === 2
        && Math.abs(segment.x0 - segment.x1) < 0.5
        && segment.x0 > RECT.w * 0.75 && segment.dash.length === 4
    )).toBe(true);
    const secondaryLabel = rec.texts.find(text =>
      text.x > RECT.w * 0.75 && text.text === '20'
    );
    expect(secondaryLabel).toMatchObject({ fillStyle: '#AABBCC' });
    expect(secondaryLabel?.font).toContain('italic 7px "Secondary Face"');
  });

  it('secondary tick-label visibility and font properties do not affect ticks or grids', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [0, 50, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'out', majorUnit: 50, tickLabelPos: 'none',
        fontBold: true, fontItalic: true, majorGridlines: true,
      },
    }), RECT, 1);
    expect(rec.texts.filter(text => text.x > RECT.w * 0.75 && /^(0|50|100)$/.test(text.text)))
      .toHaveLength(0);

    const visible = recordingCtx();
    renderChart(visible.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [0, 50, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none', majorUnit: 50,
        fontBold: true, fontItalic: true,
      },
    }), RECT, 1);
    expect(visible.texts.find(text => text.x > RECT.w * 0.75 && text.text === '50')?.font)
      .toContain('italic bold');
  });

  it('a chart with no CH6 fields renders identical gridlines to before (byte-stable)', () => {
    // Guard: the default (no CH6 fields) must keep the historical value gridlines.
    const rec = segRecordingCtx();
    renderChart(rec.ctx, lineModel({}), RECT, 1);
    expect(horizGridlines(rec.segs).length).toBeGreaterThan(2);
  });
});

// #744: `<c:catAx><c:majorGridlines>` (ECMA-376 §21.2.2.100) draws VERTICAL
// gridlines at each category tick across the plot height. The parse+type
// surface (catAxisMajorGridlines / catAxisGridlineColor / catAxisGridlineWidthEmu)
// already existed but had no renderer consumer, so a chart declaring cat-axis
// gridlines rendered without them.
describe('#744 — category-axis (vertical) major gridlines', () => {
  const colModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'clusteredBar',
    categories: ['A', 'B', 'C', 'D'],
    series: [series({ name: 'S', values: [10, 20, 30, 40] })],
    ...over,
  });

  it('OFF by default: no vertical gridlines (byte-stable)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({}), RECT, 1);
    expect(vertGridlines(rec.segs).length).toBe(0);
  });

  it('catAxisMajorGridlines=true draws vertical gridlines spanning the plot height', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisMajorGridlines: true }), RECT, 1);
    const grids = vertGridlines(rec.segs);
    // At least one gridline per category boundary/center. crossBetween="between"
    // (bar default) → n+1 dividers; either way several full-height verticals.
    expect(grids.length).toBeGreaterThanOrEqual(3);
    // Each spans (nearly) the whole plot height — much taller than a bar.
    const tallest = Math.max(...grids.map(s => Math.abs(s.y1 - s.y0)));
    expect(tallest).toBeGreaterThan(RECT.h * 0.5);
  });

  it('honors an explicit catAxisGridlineColor / width (§21.2.2.100)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, colModel({
      catAxisMajorGridlines: true,
      catAxisGridlineColor: '8fa878',
      catAxisGridlineWidthEmu: 12700, // 1 pt
    }), RECT, 1);
    const colored = vertGridlines(rec.segs, '#8fa878');
    expect(colored.length).toBeGreaterThanOrEqual(3);
    // 1 pt × ptToPx=1 → 1 px width.
    expect(colored.every(s => s.lw === 1)).toBe(true);
    // No faint default lines remain when a color is pinned.
    expect(vertGridlines(rec.segs, '#e0e0e0').length).toBe(0);
  });

  it('line chart also draws category gridlines when declared', () => {
    // The cat-gridline pass is wired into the bar, line and area renderers.
    const off = segRecordingCtx();
    renderChart(off.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, 20, 30] })],
    }), RECT, 1);
    expect(vertGridlines(off.segs).length).toBe(0);

    const on = segRecordingCtx();
    renderChart(on.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [10, 20, 30] })],
      catAxisMajorGridlines: true,
    }), RECT, 1);
    expect(vertGridlines(on.segs).length).toBeGreaterThanOrEqual(3);
  });

  it('scatter chart draws X-axis major gridlines when declared', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: [],
      series: [series({
        name: 'S',
        categories: ['0', '10', '20'],
        values: [1, 2, 3],
      })],
      catAxisMajorGridlines: true,
      catAxisGridlineColor: '8fa878',
    }), RECT, 1);
    expect(vertGridlines(rec.segs, '#8fa878').length).toBeGreaterThanOrEqual(3);
  });
});

// #738: an explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) must be honored on
// EVERY chart type's value axis, not just the primary bar/line axis. The area,
// radar and scatter renderers ignored `chart.valAxisMajorUnit`; the secondary
// (combo) axis had no majorUnit surface at all.
describe('#738 — explicit majorUnit honored on every value axis (§21.2.2.103)', () => {
  /** Numeric value-axis tick labels drawn by a chart, as numbers. */
  function valTickNumbers(over: Partial<ChartModel>): number[] {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel(over), RECT, 1);
    return rec.texts
      .map(t => t.text)
      .filter(t => /^\d+(\.\d+)?$/.test(t))
      .map(Number);
  }

  it('area: majorUnit widens the value-axis step (labels land on multiples of it)', () => {
    // Data 0..100. Auto step is fine-grained; majorUnit 50 → coarse ticks
    // 0,50,100 and NOTHING at 25/75.
    const auto = valTickNumbers({
      chartType: 'area', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [20, 60, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'area', categories: ['A', 'B', 'C'],
      series: [series({ name: 'S', values: [20, 60, 100] })],
      valAxisMajorUnit: 50,
    });
    expect(coarse).toContain(50);
    expect(coarse).not.toContain(25);
    // Coarser than auto: strictly fewer distinct tick labels.
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });

  it('scatter: majorUnit widens the Y (value) axis step', () => {
    const auto = valTickNumbers({
      chartType: 'scatter',
      series: [series({ name: 'S', values: [10, 40, 70, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'scatter',
      series: [series({ name: 'S', values: [10, 40, 70, 100] })],
      valAxisMajorUnit: 70,
    });
    expect(coarse).toContain(70);
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });

  it('radar: majorUnit widens the ring step (fewer radial ticks)', () => {
    const auto = valTickNumbers({
      chartType: 'radar', categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [20, 60, 80, 100] })],
    });
    const coarse = valTickNumbers({
      chartType: 'radar', categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [20, 60, 80, 100] })],
      valAxisMajorUnit: 70,
    });
    // Radar skips the center 0-label, so labels are the ring values.
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
    expect(coarse).toContain(70);
  });

  it('secondary (combo) axis: majorUnit widens its independent step', () => {
    // A line chart whose secondary series rides an independent right-edge axis
    // (the shared computeSecondaryAxis path, used by line/area and the bar-combo
    // line series). Secondary data 0..100; an explicit majorUnit 50 → right-side
    // ticks land on multiples of 50 (0,50,100) and NOTHING at 25.
    const secModel = (majorUnit: number | null): ChartModel => baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Big', values: [10, 20, 30] }),
        series({ name: 'Small', values: [20, 60, 100], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: null, max: null, title: 'Rate', hidden: false,
        majorTickMark: 'out', lineHidden: false, majorUnit,
      },
    });
    const rightTicks = (m: ChartModel): number[] => {
      const rec = recordingCtx();
      renderChart(rec.ctx, m, RECT, 1);
      return rec.texts
        .filter(t => t.x > RECT.x + RECT.w * 0.75 && /^\d+(\.\d+)?$/.test(t.text))
        .map(t => Number(t.text));
    };
    const auto = rightTicks(secModel(null));
    const coarse = rightTicks(secModel(50));
    expect(auto.length).toBeGreaterThan(0); // guard: right-edge ticks exist
    expect(coarse).toContain(50);
    expect(coarse).not.toContain(25);
    expect(new Set(coarse).size).toBeLessThan(new Set(auto).size);
  });
});

describe('automatic linear-axis bounds reach the shared planner', () => {
  const numericLabels = (values: number[]): number[] => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    return rec.texts.map(item => Number(item.text)).filter(Number.isFinite);
  };

  it('keeps an exact 1.2 positive line range offset and pins just above it to zero', () => {
    expect(Math.min(...numericLabels([10, 12]))).toBeGreaterThan(0);
    expect(numericLabels([10, 12.000_001])).toContain(0);
  });

  it('line charts use the shared explicit-span unit when vertical bounds omit majorUnit', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values: [0.5, 4.5] })],
      valMin: 0, valMax: 5, valAxisMajorUnit: null,
      valAxisMajorGridlines: false,
    }), RECT, 1);
    const numeric = rec.texts.map(item => Number(item.text)).filter(Number.isFinite);
    expect(numeric).toContain(0.5);
    expect(numeric).toContain(4.5);
  });

  it('line charts preserve an authored 20-unit employees axis', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      series: [
        series({ values: [42, 43, 45, 46, 48, 51, 53, 53, 56, 58, 60, 63] }),
        series({ values: [18, 18, 19, 21, 22, 22, 23, 25, 25, 26, 27, 28] }),
      ],
      valAxisMajorUnit: 20,
      valAxisMajorGridlines: false,
    }), RECT, 1);
    const ticks = rec.texts
      .filter(text => text.x < RECT.w / 2 && /^\d+$/.test(text.text))
      .map(text => Number(text.text));
    expect(ticks).toEqual([0, 20, 40, 60, 80]);
  });

  it('combo charts derive the primary employees axis from primary-bound series only', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      series: [
        series({ values: [280, 295, 310, 340, 365, 390, 410, 400, 435, 460, 490, 520] }),
        series({
          values: [0.22, 0.21, 0.23, 0.24, 0.25, 0.26, 0.26, 0.24, 0.27, 0.27, 0.28, 0.29],
          seriesType: 'line', useSecondaryAxis: true,
        }),
      ],
      valAxisMajorGridlines: false,
      secondaryValAxis: {
        min: 0, max: 0.4, title: null, hidden: false, lineHidden: false,
        majorTickMark: 'none', majorUnit: 0.05, formatCode: '0%',
      },
    }), RECT, 1);
    const primaryTicks = rec.texts
      .filter(text => text.x < RECT.w / 2 && /^\d+$/.test(text.text))
      .map(text => Number(text.text));
    expect(primaryTicks).toEqual([0, 100, 200, 300, 400, 500, 600]);
  });

  it('area geometry uses an authored non-zero minimum', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area', categories: ['Low', 'High'],
      series: [series({ values: [50, 100] })],
      valMin: 50, valMax: 100, valAxisMajorUnit: 50,
    }), RECT, 1);
    const y50 = rec.texts.find(item => item.text === '50')?.y;
    const y100 = rec.texts.find(item => item.text === '100')?.y;
    expect(y50).toBeDefined();
    expect(y100).toBeDefined();
    expect(rec.segs.some(segment =>
      Math.abs(segment.x1 - segment.x0) > 20
      && Math.abs(segment.y0 - (y50 as number)) < 1e-6
      && Math.abs(segment.y1 - (y100 as number)) < 1e-6
    )).toBe(true);
  });

  it('area negative-only data uses a negative mirrored extent for ticks and geometry', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'area', categories: ['Low', 'High'],
      series: [series({ values: [-10, -5] })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    const numeric = rec.texts.map(item => Number(item.text)).filter(Number.isFinite);
    expect(Math.min(...numeric)).toBeLessThanOrEqual(-10);
    expect(Math.max(...numeric)).toBe(0);
    expect(numeric).not.toContain(1);

    const axisTickYs = rec.texts
      .filter(item => Number.isFinite(Number(item.text)))
      .map(item => item.y);
    expect(rec.segs.some(segment =>
      Math.abs(segment.x1 - segment.x0) > 20
      && axisTickYs.some(y => Math.abs(y - segment.y0) < 1e-6)
      && axisTickYs.some(y => Math.abs(y - segment.y1) < 1e-6)
    )).toBe(true);
  });

  it('stacked area extent accumulates positive and negative values separately', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedArea', categories: ['A'],
      series: [series({ values: [10] }), series({ values: [-8] })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    const numeric = rec.texts.map(item => Number(item.text)).filter(Number.isFinite);
    expect(Math.min(...numeric)).toBeLessThanOrEqual(-8);
    expect(Math.max(...numeric)).toBeGreaterThanOrEqual(10);
  });
});

describe('authored minor tick marks are painted between major ticks', () => {
  const shortHorizontal = (segment: Seg): boolean =>
    Math.abs(segment.y1 - segment.y0) < 0.01
    && Math.abs(segment.x1 - segment.x0) > 0
    && Math.abs(segment.x1 - segment.x0) <= 12;
  const shortVertical = (segment: Seg): boolean =>
    Math.abs(segment.x1 - segment.x0) < 0.01
    && Math.abs(segment.y1 - segment.y0) > 0
    && Math.abs(segment.y1 - segment.y0) <= 12;

  it('draws primary value-axis minor ticks for a line chart', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [series({ values: [10, 50, 90] })],
      valMin: 0,
      valMax: 100,
      valAxisMajorUnit: 20,
      valAxisMinorUnit: 5,
      valAxisMajorTickMark: 'none',
      valAxisMinorTickMark: 'cross',
    }), RECT, 1);

    expect(rec.segs.filter(segment => shortHorizontal(segment) && segment.x0 < RECT.w / 2).length)
      .toBeGreaterThanOrEqual(15);
  });

  it('uses automatic major/5 positions for minor ticks without enabling minor gridlines', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ values: [10, 50, 90] })],
      valMin: 0, valMax: 100, valAxisMajorUnit: 20,
      valAxisMajorGridlines: false, valAxisMinorGridlines: false,
      valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'cross',
    }), RECT, 1);

    expect(rec.segs.filter(segment => shortHorizontal(segment) && segment.x0 < RECT.w / 2))
      .toHaveLength(20);
    expect(rec.segs.filter(segment =>
      segment.ss === '#e0e0e0' && Math.abs(segment.x1 - segment.x0) > 50
    ))
      .toHaveLength(0);
  });

  it('uses automatic major/5 positions for minor gridlines without enabling minor ticks', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [series({ values: [10, 50, 90] })],
      valMin: 0, valMax: 100, valAxisMajorUnit: 20,
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: true,
      valAxisMinorGridlineColor: '123456',
      valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    }), RECT, 1);

    const minorGridYs = rec.segs
      .filter(segment => segment.ss === '#123456' && Math.abs(segment.x1 - segment.x0) > 50)
      .map(segment => segment.y0);
    expect(minorGridYs).toHaveLength(20);
    expect(rec.segs.filter(segment =>
      shortHorizontal(segment)
      && minorGridYs.some(y => Math.abs(y - segment.y0) < 1e-6)
    )).toHaveLength(0);
  });

  it('uses automatic minor positions with horizontal value-axis geometry', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBarH', categories: ['A', 'B'],
      series: [series({ values: [20, 80] })],
      valMin: 0, valMax: 100, valAxisMajorUnit: 20,
      valAxisMajorGridlines: false,
      valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'cross',
    }), RECT, 1);

    expect(rec.segs.filter(shortVertical).length).toBeGreaterThanOrEqual(20);
  });

  it('accepts a positive authored minor unit larger than the major unit', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [series({ values: [10, 50, 90] })],
      valMin: 0,
      valMax: 100,
      valAxisMajorUnit: 20,
      valAxisMinorUnit: 30,
      valAxisMajorTickMark: 'none',
      valAxisMinorTickMark: 'cross',
    }), RECT, 1);

    // 30 and 90 are minor positions; 60 coincides with a major tick and is
    // intentionally not double-painted.
    expect(rec.segs.filter(segment => shortHorizontal(segment) && segment.x0 < RECT.w / 2))
      .toHaveLength(2);
  });

  it('anchors minor gridlines and ticks at a non-zero scale minimum', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [series({ values: [3, 13, 23] })],
      valMin: 3,
      valMax: 23,
      valAxisMajorUnit: 10,
      valAxisMinorUnit: 4,
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: true,
      valAxisMinorGridlineColor: '123456',
      valAxisMajorTickMark: 'none',
      valAxisMinorTickMark: 'cross',
    }), RECT, 1);

    const minorGridYs = rec.segs
      .filter(segment => segment.ss === '#123456' && Math.abs(segment.x1 - segment.x0) > 50)
      .map(segment => segment.y0);
    expect(minorGridYs).toHaveLength(4);
    const minorTicks = rec.segs.filter(segment =>
      shortHorizontal(segment)
      && minorGridYs.some(y => Math.abs(y - segment.y0) < 1e-6)
    );
    expect(minorTicks).toHaveLength(4);
  });

  it('bounds hostile or non-progressing minor units', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values: [1, 2] })],
      valMin: 1, valMax: 2, valAxisMajorUnit: 0.5,
      valAxisMinorUnit: Number.MIN_VALUE,
      valAxisMajorGridlines: false, valAxisMinorGridlines: true,
      valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'cross',
    }), RECT, 1);
    expect(rec.segs.length).toBeLessThan(100);

    const bounded = segRecordingCtx();
    renderChart(bounded.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values: [0, 1] })],
      valMin: 0, valMax: 1, valAxisMajorUnit: 0.5,
      valAxisMinorUnit: 1e-12,
      valAxisMajorGridlines: false, valAxisMinorGridlines: true,
      valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none',
    }), RECT, 1);
    expect(bounded.segs.length).toBeLessThanOrEqual(10_100);
  });

  it('draws a minor gridline when the positive minor unit exceeds the major unit', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values: [3, 23] })],
      valMin: 3, valMax: 23,
      valAxisMajorUnit: 10, valAxisMinorUnit: 12,
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: true,
      valAxisMinorGridlineColor: '123456',
    }), RECT, 1);
    expect(rec.segs.filter(segment =>
      segment.ss === '#123456' && Math.abs(segment.x1 - segment.x0) > 50
    )).toHaveLength(1);
  });

  it('preserves an authored major-gridline width without a solidFill color', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B'],
      series: [series({ values: [0, 20] })],
      valMin: 0, valMax: 20, valAxisMajorUnit: 10,
      valAxisMajorGridlines: true,
      valAxisGridlineWidthEmu: 25400,
      valAxisGridlineDash: 'dash',
    }), RECT, 1);
    const gridlines = rec.segs.filter(segment =>
      Math.abs(segment.x1 - segment.x0) > 50 && segment.ss === '#e0e0e0'
    );
    expect(gridlines).toHaveLength(3);
    expect(gridlines.every(segment => segment.lw === 2)).toBe(true);
  });

  it('draws minor ticks on the independent secondary value axis', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [10, 50, 90], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0,
        max: 100,
        title: null,
        hidden: false,
        majorTickMark: 'none',
        minorTickMark: 'cross',
        lineHidden: false,
        majorUnit: 20,
        minorUnit: 5,
      },
    }), RECT, 1);

    expect(rec.segs.filter(segment => shortHorizontal(segment) && segment.x0 > RECT.w / 2).length)
      .toBeGreaterThanOrEqual(15);
  });

  it('uses automatic minor positions on the independent secondary value axis', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [10, 50, 90], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false,
        majorTickMark: 'none', minorTickMark: 'cross',
        lineHidden: false, majorUnit: 20,
      },
    }), RECT, 1);

    expect(rec.segs.filter(segment => shortHorizontal(segment) && segment.x0 > RECT.w / 2))
      .toHaveLength(20);
  });

  it('keeps secondary minor gridlines and ticks independent', () => {
    const model = (minorGridlines: boolean, minorTickMark: string): ChartModel => baseModel({
      chartType: 'line', categories: ['A', 'B', 'C'],
      series: [
        series({ values: [1, 2, 3] }),
        series({ values: [10, 50, 90], useSecondaryAxis: true }),
      ],
      secondaryValAxis: {
        min: 0, max: 100, title: null, hidden: false,
        majorTickMark: 'none', minorTickMark, lineHidden: false,
        majorUnit: 20, minorGridlines,
        minorGridlineColor: '123456',
      },
    });
    const gridOnly = segRecordingCtx();
    renderChart(gridOnly.ctx, model(true, 'none'), RECT, 1);
    expect(gridOnly.segs.filter(segment =>
      segment.ss === '#123456' && Math.abs(segment.x1 - segment.x0) > 50
    )).toHaveLength(20);
    expect(gridOnly.segs.filter(segment => shortHorizontal(segment) && segment.x0 > RECT.w / 2))
      .toHaveLength(0);

    const tickOnly = segRecordingCtx();
    renderChart(tickOnly.ctx, model(false, 'cross'), RECT, 1);
    expect(tickOnly.segs.filter(segment => segment.ss === '#123456')).toHaveLength(0);
    expect(tickOnly.segs.filter(segment => shortHorizontal(segment) && segment.x0 > RECT.w / 2))
      .toHaveLength(20);
  });

  it('draws both numeric X- and Y-axis minor ticks for scatter', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      series: [series({ values: [10, 50, 90], categories: ['10', '50', '90'] })],
      valMin: 0,
      valMax: 100,
      valAxisMajorUnit: 20,
      valAxisMinorUnit: 5,
      catAxisMajorUnit: 20,
      catAxisMinorUnit: 5,
      valAxisMajorTickMark: 'none',
      catAxisMajorTickMark: 'none',
      valAxisMinorTickMark: 'cross',
      catAxisMinorTickMark: 'cross',
    }), RECT, 1);

    expect(rec.segs.filter(shortHorizontal).length).toBeGreaterThanOrEqual(15);
    expect(rec.segs.filter(shortVertical).length).toBeGreaterThanOrEqual(15);
  });
});

describe('radar value-axis planning', () => {
  it('honors an explicit minimum and paints minor gridlines without minor ticks', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'radar', radarStyle: 'standard',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({ values: [50, 60, 70, 80] })],
      valMin: 50, valMax: 100, valAxisMajorUnit: 25,
      valAxisMinorGridlines: true, valAxisMinorGridlineColor: '123456',
      valAxisMinorTickMark: 'none',
    }), RECT, 1);
    expect(rec.texts.some(item => item.text === '50')).toBe(false);
    expect(rec.texts.some(item => item.text === '75')).toBe(true);
    expect(rec.segs.filter(segment => segment.ss === '#123456').length).toBeGreaterThan(0);
  });

  it('paints radar minor ticks without enabling minor gridlines', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'radar', radarStyle: 'standard',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({ values: [50, 60, 70, 80] })],
      valMin: 50, valMax: 100, valAxisMajorUnit: 25,
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: false,
      valAxisMinorTickMark: 'cross',
    }), RECT, 1);
    expect(rec.segs.filter(segment => segment.ss === '#e0e0e0')).toHaveLength(0);
    expect(rec.segs.filter(segment =>
      Math.abs(segment.y1 - segment.y0) < 1e-6
      && Math.abs(segment.x1 - segment.x0) > 0
      && Math.abs(segment.x1 - segment.x0) <= 12
    ).length).toBeGreaterThan(0);
  });

  it('keeps authored radar grid, axis-label and category-label styles', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'radar', radarStyle: 'standard',
      categories: ['North', 'East', 'South', 'West'],
      series: [series({ values: [1, 2, 3, 4] })],
      valMin: 0, valMax: 4, valAxisMajorUnit: 2,
      valAxisMajorGridlines: true,
      valAxisGridlineColor: '123456',
      valAxisGridlineWidthEmu: 25_400,
      valAxisFontFace: 'Value Face',
      valAxisFontSizeHpt: 1_100,
      valAxisFontBold: true,
      valAxisFontColor: '654321',
      catAxisFontFace: 'Category Face',
      catAxisFontSizeHpt: 900,
      catAxisFontItalic: true,
      catAxisFontColor: 'ABCDEF',
    }), RECT, 1);

    const valueLabel = rec.texts.find(item => item.text === '2');
    expect(valueLabel?.font).toContain('bold');
    expect(valueLabel?.font).toContain('11px "Value Face"');
    expect(valueLabel?.fillStyle).toBe('#654321');
    const categoryLabel = rec.texts.find(item => item.text === 'North');
    expect(categoryLabel?.font).toContain('italic');
    expect(categoryLabel?.font).toContain('9px "Category Face"');
    expect(categoryLabel?.fillStyle).toBe('#ABCDEF');
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#123456'
    )).toBe(true);
  });

  it('keeps radar series line/noFill and indexed marker overrides authoritative', () => {
    const hidden = strokedPolylineCtx();
    renderChart(hidden.ctx, baseModel({
      chartType: 'radar', radarStyle: 'marker',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({
        values: [1, 2, 3, 4],
        lineHidden: true,
        lineColor: '135790',
        lineWidthEmu: 38_100,
        showMarker: false,
      })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    expect(hidden.strokes.some(stroke => stroke.ss === '#135790')).toBe(false);

    const styled = strokedPolylineCtx();
    renderChart(styled.ctx, baseModel({
      chartType: 'radar', radarStyle: 'standard',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({
        values: [1, 2, 3, 4],
        lineColor: '2468AC',
        lineWidthEmu: 25_400,
        chartexStyle: { lineDash: 'dash', lineCap: 'rnd', lineJoin: 'bevel' },
        showMarker: false,
      })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    expect(styled.strokes).toContainEqual(expect.objectContaining({
      ss: '#2468AC', lw: 2, dash: [12, 6], cap: 'round', join: 'bevel',
    }));

    const markers = recordingCtx();
    renderChart(markers.ctx, baseModel({
      chartType: 'radar', radarStyle: 'marker',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({
        values: [1, 2, 3, 4],
        lineHidden: true,
        markerSymbol: 'circle', markerSize: 4,
        dataPointOverrides: [{ idx: 1, markerSymbol: 'circle', markerSize: 10 }],
      })],
      valAxisMajorGridlines: false,
    }), RECT, 1);
    expect(markers.arcs.some(arc => Math.abs(arc.r - 5) < 1e-6)).toBe(true);
    expect(markers.arcs.filter(arc => Math.abs(arc.r - 2) < 1e-6).length).toBeGreaterThan(0);
  });
});

/** Recording context that counts rotate() calls and captures fillText, for the
 *  category-label rotation / tickLblPos tests. */
function rotateRecordingCtx(): { ctx: CanvasRenderingContext2D; rotates: number[]; texts: string[] } {
  const rotates: number[] = [];
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const fontPx = (font: string): number => {
    const m = /(\d+(?:\.\d+)?)px/.exec(font);
    return m ? parseFloat(m[1]) : 10;
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'rotate': return (r: number) => { rotates.push(r); };
        case 'fillText': return (text: string) => texts.push(String(text));
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'fill': case 'stroke': case 'moveTo': case 'lineTo': case 'arc':
        case 'bezierCurveTo': case 'quadraticCurveTo': case 'rect': case 'fillRect':
        case 'strokeRect': case 'clearRect': case 'strokeText': case 'setLineDash':
        case 'translate': case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, rotates, texts };
}

describe('CH6 — category-axis label rotation + tickLblPos (commit 2)', () => {
  const colModel = (over: Partial<ChartModel>): ChartModel => baseModel({
    chartType: 'clusteredBar',
    categories: ['Alpha', 'Beta', 'Gamma'],
    series: [series({ name: 'S', values: [10, 20, 30] })],
    ...over,
  });

  it('catAxisTickLabelPos="none" hides the category labels', () => {
    const shown = rotateRecordingCtx();
    renderChart(shown.ctx, colModel({}), RECT, 1);
    expect(shown.texts.some(t => t.startsWith('Alpha'))).toBe(true);

    const hidden = rotateRecordingCtx();
    renderChart(hidden.ctx, colModel({ catAxisTickLabelPos: 'none' }), RECT, 1);
    expect(hidden.texts.some(t => t.startsWith('Alpha'))).toBe(false);
    // Value tick labels still present.
    expect(hidden.texts.some(t => /^\d+$/.test(t))).toBe(true);
  });

  it('catAxisLabelRotation rotates the column category labels', () => {
    const flat = rotateRecordingCtx();
    renderChart(flat.ctx, colModel({}), RECT, 1);
    expect(flat.rotates.length).toBe(0);

    const rot = rotateRecordingCtx();
    // -2700000 60000ths = -45°.
    renderChart(rot.ctx, colModel({ catAxisLabelRotation: -2_700_000 }), RECT, 1);
    expect(rot.rotates.length).toBeGreaterThan(0);
    const rad = rot.rotates[0];
    expect(rad).toBeCloseTo((-45 * Math.PI) / 180, 6);
    // Labels still drawn (just rotated).
    expect(rot.texts.some(t => t.startsWith('Alpha'))).toBe(true);
  });

  it('rotation 0 keeps the un-rotated fast path (byte-stable, no rotate calls)', () => {
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: 0 }), RECT, 1);
    expect(rec.rotates.length).toBe(0);
  });

  it('honors lblAlgn inside each column category interval', () => {
    const label = (alignment: 'l' | 'ctr' | 'r') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, colModel({ catAxisLabelAlignment: alignment }), RECT, 1);
      return rec.texts.find(text => text.text === 'Alpha')!;
    };
    const left = label('l');
    const center = label('ctr');
    const right = label('r');
    expect(left.align).toBe('left');
    expect(center.align).toBe('center');
    expect(right.align).toBe('right');
    expect(left.x).toBeLessThan(center.x);
    expect(center.x).toBeLessThan(right.x);
  });

  it('honors lblAlgn in the horizontal-bar category-label gutter', () => {
    const label = (alignment: 'l' | 'ctr' | 'r') => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBarH',
        categories: ['Alpha', 'Beta'],
        series: [series({ name: 'S', values: [10, 20] })],
        catAxisLabelAlignment: alignment,
      }), RECT, 1);
      return rec.texts.find(text => text.text === 'Alpha')!;
    };
    const left = label('l');
    const center = label('ctr');
    const right = label('r');
    expect(left.align).toBe('left');
    expect(center.align).toBe('center');
    expect(right.align).toBe('right');
    expect(left.x).toBeLessThan(center.x);
    expect(center.x).toBeLessThan(right.x);
  });

  it('scales the column rule-to-label gap from the established default', () => {
    const measure = (offset: number) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, colModel({
        catAxisFontSizeHpt: 900,
        catAxisLabelOffsetPercent: offset,
        valMin: 0,
        valMax: 40,
      }), RECT, 1);
      const label = rec.texts.find(text => text.text.startsWith('Al'))!;
      const baseline = Math.max(...rec.rects.map(rect => rect.y + rect.h));
      return label.y - baseline;
    };
    expect(measure(0)).toBeCloseTo(0, 6);
    expect(measure(250)).toBeCloseTo(measure(100) * 2.5, 6);
  });

  it('scales the horizontal-bar rule-to-label gap from the established default', () => {
    const measure = (offset: number) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'clusteredBarH',
        categories: ['Alpha', 'Beta'],
        series: [series({ name: 'S', values: [10, 20] })],
        catAxisFontSizeHpt: 900,
        catAxisLabelAlignment: 'r',
        catAxisLabelOffsetPercent: offset,
        valMin: 0,
        valMax: 30,
      }), RECT, 1);
      const label = rec.texts.find(text => text.text.startsWith('Al'))!;
      const plotLeft = Math.min(...rec.rects.map(rect => rect.x));
      return plotLeft - label.x;
    };
    expect(measure(0)).toBeCloseTo(0, 6);
    expect(measure(250)).toBeCloseTo(measure(100) * 2.5, 6);
  });

  it('wraps long horizontal category labels without discarding words', () => {
    const longLabel = 'Foundations: Economic growth and inclusive development';
    const rec = recordingCtx();
    renderChart(rec.ctx, colModel({
      categories: [longLabel, 'Priority 1', 'Priority 2', 'Priority 3', 'Priority 4', 'Applications'],
      series: [series({ name: 'S', values: [1, 2, 3, 4, 5, 6] })],
      catAxisFontSizeHpt: 900,
      catAxisLabelRotation: -60_000_000,
    }), RECT, 1);

    const lines = rec.texts.filter(text => longLabel.includes(text.text));
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.map(line => line.text).join(' ')).toBe(longLabel);
    expect(lines.some(line => line.text.includes('…'))).toBe(false);
    expect(lines.map(line => line.y)).toEqual([...lines.map(line => line.y)].sort((a, b) => a - b));
  });

  // #748: a rot outside the ST_FixedAngle (§20.1.10.23) (-90°,90°) text-rotation
  // range is not a valid axis-label rotation — Office draws such labels
  // horizontal. sample-24's cat/date/value axes all carry rot="-60000000"
  // (-1000°) yet Word renders every label horizontal (verified against
  // sample-24.pdf: "Category" label bbox is wide/short, ratio ≈ 3.0). Naively
  // dividing -60000000/60000 = -1000° (or wrapping mod 360 → +80°) rotates them
  // near-vertical, which is wrong.
  it('an out-of-range rot ("-60000000" = -1000°) draws labels HORIZONTAL, not rotated', () => {
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: -60_000_000 }), RECT, 1);
    // Office ignores the out-of-range rotation: no rotate() calls (horizontal
    // fast path), labels still drawn.
    expect(rec.rotates.length).toBe(0);
    expect(rec.texts.some(t => t.startsWith('Alpha'))).toBe(true);
  });

  it('a rot at the ±90° ST_FixedAngle boundary (-5400000 = -90°) is still honored', () => {
    // -90° is the inclusive edge of Office's axis-text rotation range; keep it
    // working (genuine vertical axis labels).
    const rec = rotateRecordingCtx();
    renderChart(rec.ctx, colModel({ catAxisLabelRotation: -5_400_000 }), RECT, 1);
    expect(rec.rotates.length).toBeGreaterThan(0);
    expect(rec.rotates[0]).toBeCloseTo((-90 * Math.PI) / 180, 6);
  });
});

/** Recording context that captures line-dash state alongside stroked segments. */
function dashSegRecordingCtx(): { ctx: CanvasRenderingContext2D; segs: Array<{ dashed: boolean }> } {
  const segs: Array<{ dashed: boolean }> = [];
  let dash: number[] = [];
  let pending = false;
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (t: string) => ({ width: String(t).length * 6 });
        case 'setLineDash': return (d: number[]) => { dash = d ?? []; };
        case 'getLineDash': return () => dash;
        case 'lineTo': return () => { pending = true; };
        case 'stroke': return () => { if (pending) { segs.push({ dashed: dash.length > 0 }); pending = false; } };
        case 'createLinearGradient': case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        case 'save': case 'restore': case 'beginPath': case 'closePath':
        case 'fill': case 'moveTo': case 'arc': case 'bezierCurveTo':
        case 'quadraticCurveTo': case 'rect': case 'fillRect': case 'strokeRect':
        case 'clearRect': case 'fillText': case 'strokeText': case 'translate':
        case 'rotate': case 'scale': case 'clip': case 'setTransform':
        case 'resetTransform': case 'getTransform':
          return () => undefined;
        default: return undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, segs };
}

describe('CH6-follow — series trendlines (commit 3)', () => {
  const lineWithTrend = (over: Partial<ChartSeries>): ChartModel => baseModel({
    chartType: 'line',
    categories: ['A', 'B', 'C', 'D'],
    series: [series({ name: 'S', values: [1, 3, 5, 7], ...over })],
  });

  it('a linear trendline without prstDash draws an additional solid line', () => {
    const noTrend = dashSegRecordingCtx();
    renderChart(noTrend.ctx, lineWithTrend({}), RECT, 1);
    expect(noTrend.segs.some(s => s.dashed)).toBe(false);

    const withTrend = dashSegRecordingCtx();
    renderChart(withTrend.ctx, lineWithTrend({ trendLines: [{ trendlineType: 'linear' }] }), RECT, 1);
    expect(withTrend.segs.length).toBeGreaterThan(noTrend.segs.length);
    expect(withTrend.segs.some(s => s.dashed)).toBe(false);
  });

  it('renders column-series percentage error bars, trendline equation, and legend key', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      showLegend: true,
      legendPos: 'b',
      valMin: 0,
      valMax: 200,
      series: [series({
        name: 'Monthly Revenue',
        color: '4472C4',
        values: [100, 120, 110, 150, 140, 170],
        errBars: [{
          dir: 'y', barType: 'both', noEndCap: true,
          plus: [10, 12, 11, 15, 14, 17],
          minus: [10, 12, 11, 15, 14, 17],
          color: '404040',
        }],
        trendLines: [{
          trendlineType: 'linear', dispEq: true, lineColor: 'FF0000',
        }],
      })],
    }), RECT, 1);

    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#404040' });
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF0000' });
    expect(rec.texts.map(text => text.text)).toContain('y = 12.8571x + 86.6667');
    expect(rec.texts.map(text => text.text)).toContain('Monthly Revenue');
    expect(rec.texts.map(text => text.text)).toContain('Linear (Monthly Revenue)');
  });

  it('includes column error-bar endpoints in the automatic value-axis extent', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [series({
        values: [100, 120, 110, 150, 140, 170],
        errBars: [{
          dir: 'y', barType: 'both', noEndCap: true,
          plus: [10, 12, 11, 15, 14, 17],
          minus: [10, 12, 11, 15, 14, 17],
        }],
      })],
    }), RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toContain('0');
    expect(labels).toContain('200');
    expect(labels).not.toContain('-20');
  });

  it('uses an authored trendline name in the bar legend', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredColumn',
      categories: ['A', 'B'],
      showLegend: true,
      legendPos: 'b',
      series: [series({
        name: 'Series', values: [1, 2],
        trendLines: [{ trendlineType: 'linear', name: 'Forecast' }],
      })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toContain('Forecast');
  });

  it('adds a line-series trendline to the shared legend path', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({ trendLines: [{ trendlineType: 'linear' }] }),
      showLegend: true,
      legendPos: 'b',
    }, RECT, 1);
    expect(rec.texts.map(text => text.text)).toContain('Linear (S)');
  });

  it('honors the trendline DrawingML dash preset', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{ trendlineType: 'linear', lineDash: 'dash' }],
    }), RECT, 1);
    expect(rec.segs.some(segment => segment.dashed)).toBe(true);
  });

  it('uses the linked trendline role behind omitted trendline line properties', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({ trendLines: [{ trendlineType: 'linear' }] }),
      showLegend: true,
      legendPos: 'b',
      chartStyleRoles: {
        trendline: { lineColors: ['AABBCC'], lineWidthEmu: 19_050, lineDash: 'dash' },
      },
    }, RECT, 1);
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#AABBCC' });
    expect(rec.texts.map(text => text.text)).toContain('Linear (S)');
  });

  it('keeps direct trendline paint ahead of the linked trendline role', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        trendLines: [{ trendlineType: 'linear', lineColor: '112233', lineWidthEmu: 38_100 }],
      }),
      chartStyleRoles: {
        trendline: { lineColors: ['AABBCC'], lineWidthEmu: 19_050 },
      },
    }, RECT, 1);
    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#112233' });
    expect(rec.paintEvents).not.toContainEqual({ kind: 'stroke', strokeStyle: '#AABBCC' });
  });

  it('applies linked trendline-label text body and structured shape before layout', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        trendLines: [{ trendlineType: 'linear', dispEq: true }],
      }),
      chartStyleRoles: {
        trendlineLabel: {
          fontSizeHpt: 1600,
          fontBold: true,
          fontItalic: true,
          fontColor: '123456',
          textRotation: 5_400_000,
          textLInsEmu: 12_700,
          textRInsEmu: 12_700,
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0, stops: [
              { position: 0, color: 'FF0000' },
              { position: 1, color: '0000FF' },
            ],
          }],
          lineColors: ['445566'],
          lineWidthEmu: 25_400,
        },
      },
    }, RECT, 1);

    const equation = rec.texts.find(text => text.text.startsWith('y = '));
    expect(equation?.font).toContain('italic bold 16px');
    expect(equation?.fillStyle).toBe('#123456');
    expect(rec.rotations).toContainEqual(Math.PI / 2);
    expect(rec.gradients).toHaveLength(1);
    expect(rec.strokeRects).toContainEqual(expect.objectContaining({ ss: '#445566', lw: 2 }));
  });

  it('uses manual trendline-label geometry for the shape and preserves no-wrap text', () => {
    const rec = recordingCtx();
    const labelText = 'This authored label remains on one line';
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{
        trendlineType: 'linear', dispEq: true, labelText,
        labelTextRotation: 5_400_000,
        labelTextWrap: 'none',
        labelManualLayout: {
          xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
          x: 0.1, y: 0.2, w: 0.25, h: 0.1,
        },
        labelBox: { fill: 'FF00FF' },
      }],
    }), RECT, 1);
    expect(rec.rects).toContainEqual(expect.objectContaining({
      w: RECT.w * 0.25,
      h: RECT.h * 0.1,
      fs: '#FF00FF',
    }));
    expect(rec.rotations).toContainEqual(Math.PI / 2);
    expect(rec.texts.some(text => text.text === labelText)).toBe(true);
    expect(rec.texts.some(text => text.text.includes('…'))).toBe(false);
  });

  it('preserves trendline-label rich runs and paragraph-specific alignment', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{
        trendlineType: 'linear', dispEq: true,
        labelText: 'LongLeft\nR',
        labelRichRuns: [
          { text: 'LongLeft', color: '112233', paragraphAlign: 'l' },
          { text: '\n' },
          { text: 'R', color: '445566', italic: true, paragraphAlign: 'r' },
        ],
      }],
    }), RECT, 1);
    const left = rec.texts.find(text => text.text === 'LongLeft');
    const right = rec.texts.find(text => text.text === 'R');
    expect(left?.fillStyle).toBe('#112233');
    expect(right?.fillStyle).toBe('#445566');
    expect(right?.font).toContain('italic');
    expect((right?.x ?? 0)).toBeGreaterThan(left?.x ?? 0);
  });

  it('keeps one hidden trendline-label run from hiding its unformatted sibling', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        trendLines: [{
          trendlineType: 'linear', dispEq: true, labelText: 'HIDDENVISIBLE',
          labelRichRuns: [
            { text: 'HIDDEN', colorPaintAuthored: true, colorHidden: true },
            { text: 'VISIBLE' },
          ],
        }],
      }),
      chartStyleRoles: { trendlineLabel: { fontColor: '008000' } },
    }, RECT, 1);
    expect(rec.texts.some(text => text.text === 'HIDDEN')).toBe(false);
    expect(rec.texts).toContainEqual(expect.objectContaining({ text: 'VISIBLE', fillStyle: '#008000' }));
  });

  it('keeps direct trendline-label text and no-fill shape ahead of the linked role', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        trendLines: [{
          trendlineType: 'linear', dispEq: true,
          labelFontSizeHpt: 900,
          labelFontBold: false,
          labelFontItalic: false,
          labelFontColor: '654321',
          labelTextRotation: 0,
          labelBox: { fillHidden: true, fillPaintAuthored: true, borderHidden: true },
        }],
      }),
      chartStyleRoles: {
        trendlineLabel: {
          fontSizeHpt: 1600, fontBold: true, fontItalic: true, fontColor: '123456',
          textRotation: 5_400_000,
          fillPaints: [{ fillType: 'solid', color: 'FF0000' }],
          lineColors: ['445566'],
        },
      },
    }, RECT, 1);

    const equation = rec.texts.find(text => text.text.startsWith('y = '));
    expect(equation?.font).toContain('9px');
    expect(equation?.font).not.toContain('italic');
    expect(equation?.fillStyle).toBe('#654321');
    expect(rec.rotations).not.toContainEqual(Math.PI / 2);
    expect(rec.rects.every(rect => rect.fs !== '#FF0000')).toBe(true);
    expect(rec.strokeRects.every(rect => rect.ss !== '#445566')).toBe(true);
  });

  it('applies linked data-label run and body defaults without overriding the series', () => {
    const linked = recordingCtx();
    renderChart(linked.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
        },
      })],
      chartStyleRoles: {
        dataLabel: {
          fontSizeHpt: 1500, fontBold: true, fontItalic: true, fontColor: 'AABBCC',
          textRotation: 2_700_000,
        },
      },
    }), RECT, 1);
    const value = linked.texts.find(text => text.text === '1' && text.fillStyle === '#AABBCC');
    expect(value?.font).toContain('italic bold 15px');
    expect(value?.fillStyle).toBe('#AABBCC');
    expect(linked.rotations).toContainEqual(Math.PI / 4);

    const direct = recordingCtx();
    renderChart(direct.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          fontSizeHpt: 800, fontBold: false, fontItalic: false, fontColor: '112233',
          textRotation: 0,
        },
      })],
      chartStyleRoles: {
        dataLabel: {
          fontSizeHpt: 1500, fontBold: true, fontItalic: true, fontColor: 'AABBCC',
          textRotation: 2_700_000,
        },
      },
    }), RECT, 1);
    const directValue = direct.texts.find(text => text.text === '1' && text.fillStyle === '#112233');
    expect(directValue?.font).toContain('8px');
    expect(directValue?.font).not.toContain('italic');
    expect(directValue?.fillStyle).toBe('#112233');
    expect(direct.rotations).not.toContainEqual(Math.PI / 4);
  });

  it('uses linked data-label body formatting in the optional 3-D renderer', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
        },
      })],
      chartStyleRoles: {
        dataLabel: {
          fontSizeHpt: 1400,
          fontItalic: true,
          fontColor: 'AABBCC',
          fontBaseline: 0.25,
          textRotation: 5_400_000,
          textLInsEmu: 12_700,
        },
      },
    }), RECT, 1);

    const value = rec.texts.find(text => text.text === '1' && text.fillStyle === '#AABBCC');
    expect(value?.font).toContain('italic 14px');
    expect(rec.rotations).toContainEqual(Math.PI / 2);
  });

  it('keeps optional 3-D edge labels inside the clamped plot after alignment resolution', () => {
    const rec = recordingCtx();
    const label = 'A deliberately wide automatic label near the projected plot edge';
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A'],
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      series: [series({
        values: [1],
        dataLabelOverrides: [{ idx: 0, text: label, position: 'r' }],
      })],
    }), RECT, 1);

    const value = rec.texts.find(text => text.text.startsWith('A deliberately'));
    expect(value?.align).toBe('left');
    const labelClip = rec.clips
      .filter(clip => value != null && value.y >= clip.y && value.y <= clip.y + clip.h)
      .sort((a, b) => a.w * a.h - b.w * b.h)[0];
    expect(labelClip).toBeDefined();
    expect((value as TextCall).x).toBeGreaterThanOrEqual((labelClip as { x: number }).x - 1e-6);
  });

  it('applies linked callout paint and body formatting to pie data labels', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['A', 'B'],
      series: [series({
        values: [1, 1],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: false,
          labelBox: {},
        },
      })],
      chartStyleRoles: {
        dataLabelCallout: {
          fontItalic: true,
          fontColor: 'AABBCC',
          textRotation: 2_700_000,
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0, stops: [
              { position: 0, color: 'FF0000' },
              { position: 1, color: '0000FF' },
            ],
          }],
          lineColors: ['445566'],
          lineWidthEmu: 25_400,
        },
      },
    }), RECT, 1);

    const labels = rec.texts.filter(text => text.text === 'A' || text.text === 'B');
    expect(labels.length).toBeGreaterThan(0);
    expect(labels.every(text => text.font?.includes('italic'))).toBe(true);
    expect(rec.rotations).toContainEqual(Math.PI / 4);
    expect(rec.gradients.length).toBeGreaterThan(0);
    expect(rec.strokeRects).toContainEqual(expect.objectContaining({ ss: '#445566', lw: 2 }));
  });

  it('does not turn an ordinary indexed data label into a linked callout', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'waterfall',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
        },
        dataLabelOverrides: [{ idx: 0, text: '1' }],
      })],
      chartStyleRoles: {
        dataLabelCallout: {
          fillPaints: [{
            fillType: 'gradient', gradType: 'linear', angle: 0, stops: [
              { position: 0, color: 'FF0000' },
              { position: 1, color: '0000FF' },
            ],
          }],
          lineColors: ['445566'],
          lineWidthEmu: 25_400,
        },
      },
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === '1')).toBe(true);
    expect(rec.gradients).toHaveLength(0);
    expect(rec.strokeRects.some(rect => rect.ss === '#445566')).toBe(false);
  });

  it('merges point, series, and linked callout properties without reviving text noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: false,
          labelBox: { fill: '224466' },
        },
        dataLabelOverrides: [{
          idx: 0,
          text: '',
          fontPaintAuthored: true,
          fontHidden: true,
          labelBox: { borderWidthEmu: 25_400 },
        }],
      })],
      chartStyleRoles: {
        dataLabelCallout: {
          fontColor: 'AABBCC',
          fillColors: ['00FF00'],
          lineColors: ['FF0000'],
        },
      },
    }), RECT, 1);
    expect(rec.rects.some(rect => rect.fs === '#224466')).toBe(true);
    expect(rec.strokeRects).toContainEqual(expect.objectContaining({ ss: '#FF0000', lw: 2 }));
    expect(rec.texts.some(text => text.text === 'A')).toBe(false);
  });

  it('rotates a legend key and its linked data-label text as one measured block', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A'],
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: true, showCatName: false, showSerName: false, showPercent: false,
          showLegendKey: true,
        },
      })],
      chartStyleRoles: {
        dataLabel: { fontItalic: true, textRotation: 5_400_000 },
      },
    }), RECT, 1);

    expect(rec.texts.some(text => text.text === '1' && text.font?.includes('italic'))).toBe(true);
    expect(rec.rotations).toContainEqual(Math.PI / 2);
  });

  it('suppresses a trendline when the linked trendline role declares noFill', () => {
    const without = dashSegRecordingCtx();
    renderChart(without.ctx, lineWithTrend({}), RECT, 1);
    const hidden = dashSegRecordingCtx();
    renderChart(hidden.ctx, {
      ...lineWithTrend({ trendLines: [{ trendlineType: 'linear' }] }),
      chartStyleRoles: { trendline: { lineHidden: true } },
    }, RECT, 1);
    expect(hidden.segs).toEqual(without.segs);
  });

  it('suppresses a trendline with an authored DrawingML noFill line', () => {
    const without = dashSegRecordingCtx();
    renderChart(without.ctx, lineWithTrend({}), RECT, 1);
    const hidden = dashSegRecordingCtx();
    renderChart(hidden.ctx, lineWithTrend({
      trendLines: [{ trendlineType: 'linear', lineHidden: true }],
    }), RECT, 1);
    expect(hidden.segs).toEqual(without.segs);
  });

  it('aligns automatic equation blocks to one plot-relative column while following endpoint height', () => {
    const renderLabels = (values: number[]) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, lineWithTrend({
        values,
        trendLines: [{ trendlineType: 'linear', dispEq: true, dispRSqr: true }],
      }), RECT, 1);
      return rec.texts.filter(text => text.text.startsWith('y = ') || text.text.startsWith('R² = '));
    };
    const positive = renderLabels([1, 3, 5, 7]);
    const negative = renderLabels([7, 5, 3, 1]);
    expect(positive[0].x + (positive[0].width ?? 0)).toBe(
      negative[0].x + (negative[0].width ?? 0),
    );
    expect(positive.map(text => text.y)).not.toEqual(negative.map(text => text.y));
    expect(positive.map(text => text.text)).toEqual(['y = 2x + 1', 'R² = 1']);
    expect(positive).toHaveLength(2);
    for (const text of [...positive, ...negative]) {
      expect(text.x).toBeGreaterThanOrEqual(RECT.x);
      expect(text.x + (text.width ?? 0)).toBeLessThanOrEqual(RECT.x + RECT.w);
      expect(text.y).toBeGreaterThanOrEqual(RECT.y);
      expect(text.y).toBeLessThanOrEqual(RECT.y + RECT.h);
    }
  });

  it('keeps both automatic equation and R-squared lines with an empty authored bodyPr', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      values: [3, 3, 5, 6],
      trendLines: [{
        trendlineType: 'linear',
        dispEq: true,
        dispRSqr: true,
        labelTextBodyAuthored: true,
        labelFontSizeHpt: 800,
        labelBox: { fill: 'FFFFFF', borderColor: '808080', borderWidthEmu: 3175 },
      }],
    }), RECT, 4 / 3);

    expect(rec.texts.filter(text => text.text.startsWith('y = '))).toHaveLength(1);
    expect(rec.texts.filter(text => text.text.startsWith('R² = '))).toHaveLength(1);
  });

  it('honors trendline-label manual layout and text properties independently of line paint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{
        trendlineType: 'linear',
        lineHidden: true,
        dispEq: true,
        labelText: 'Authored fit',
        labelFontSizeHpt: 1800,
        labelFontBold: true,
        labelFontColor: '123456',
        labelFontFace: 'Georgia',
        labelTextAlign: 'ctr',
        labelManualLayout: {
          xMode: 'edge', yMode: 'edge', wMode: 'factor', hMode: 'factor',
          x: 0.1, y: 0.2, w: 0.25, h: 0.1,
        },
      }],
    }), RECT, 1);
    const label = rec.texts.find(text => text.text === 'Authored fit');
    expect(label).toMatchObject({
      x: RECT.x + RECT.w * 0.225,
      y: RECT.y + RECT.h * 0.2,
      align: 'center',
      baseline: 'top',
      fillStyle: '#123456',
    });
    expect(label?.font).toContain('bold 18px');
    expect(label?.font).toContain('Georgia');
  });

  it('uses shared data-label bold and color when trendline text properties are omitted', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        trendLines: [{ trendlineType: 'linear', labelText: 'First\nSecond' }],
      }),
      dataLabelFontBold: true,
      dataLabelFontColor: 'FF0000',
    }, RECT, 1);
    const labels = rec.texts.filter(text => text.text === 'First' || text.text === 'Second');
    expect(labels.map(label => label.text)).toEqual(['First', 'Second']);
    expect(labels.every(label => label.font?.includes('bold'))).toBe(true);
    expect(labels.every(label => label.fillStyle === '#FF0000')).toBe(true);
  });

  it.each([
    { flags: { dispEq: true }, expected: ['y = 2x + 1'] },
    { flags: { dispRSqr: true }, expected: ['R² = 1'] },
  ])('renders $expected without coupling text calculation to placement', ({ flags, expected }) => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{ trendlineType: 'linear', ...flags }],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text).filter(text => text.startsWith('y =') || text.startsWith('R²')))
      .toEqual(expected);
  });

  it('formats generated trendline values with the authored label numFmt', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{ trendlineType: 'linear', dispEq: true, dispRSqr: true, labelFormatCode: '0.00' }],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining([
      'y = 2.00x + 1.00',
      'R² = 1.00',
    ]));
  });

  it('uses the source-series numFmt for a source-linked trendline label', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      valFormatCode: '0.0%',
      trendLines: [{
        trendlineType: 'linear', dispEq: true, dispRSqr: true,
        labelFormatCode: '0.00', labelFormatSourceLinked: true,
      }],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining([
      'y = 200.0%x + 100.0%',
      'R² = 100.0%',
    ]));
  });

  it('uses the rightmost fitted value only for automatic label height', () => {
    const labelY = (values: number[]): number => {
      const rec = recordingCtx();
      renderChart(rec.ctx, lineWithTrend({
        values,
        trendLines: [{ trendlineType: 'linear', dispEq: true }],
      }), RECT, 1);
      const label = rec.texts.find(text => text.text.startsWith('y = '));
      return label?.y ?? 0;
    };
    expect(labelY([1, 1.01, 1.02, 1.03])).not.toBe(labelY([1, 51, 101, 151]));
  });

  it('paints authored trendline-label fill, border, italic text, and width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{
        trendlineType: 'linear',
        labelText: 'boxed fit',
        labelFontItalic: true,
        labelBox: { fill: 'FFFFFF', borderColor: '808080', borderWidthEmu: 25_400 },
      }],
    }), RECT, 1);
    expect(rec.rects.some(rect => rect.fs === '#FFFFFF')).toBe(true);
    expect(rec.strokeRects.some(rect => rect.ss === '#808080' && rect.lw === 2)).toBe(true);
    expect(rec.texts.find(text => text.text === 'boxed fit')?.font).toContain('italic');
  });

  it('does not send overflowed regression coordinates or labels to Canvas', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, {
      ...lineWithTrend({
        values: [Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE],
        trendLines: [{ trendlineType: 'linear', dispEq: true, dispRSqr: true }],
      }),
      valMin: 0,
      valMax: 10,
    }, RECT, 1);
    expect(rec.texts.map(text => text.text).join(' ')).not.toMatch(/NaN|Infinity/);
    expect(rec.texts.some(text => text.text.startsWith('y = ') || text.text.startsWith('R² = ')))
      .toBe(false);
  });

  it('rejects non-finite trendline geometry after finite forward extension', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, lineWithTrend({
      trendLines: [{
        trendlineType: 'linear',
        forward: Number.MAX_VALUE,
        dispEq: true,
      }],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text).join(' ')).not.toMatch(/NaN|Infinity/);
    expect(rec.texts.some(text => text.text.startsWith('y = '))).toBe(false);
  });

  it('renders a scatter trendline label through the shared plot-aware path', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'scatter',
      categories: [],
      series: [series({
        categories: ['0', '1', '2', '3'],
        values: [1, 3, 5, 7],
        trendLines: [{ trendlineType: 'linear', dispEq: true }],
      })],
      catAxisMin: 0,
      catAxisMax: 3,
      valMin: 0,
      valMax: 8,
    }), RECT, 1);
    expect(rec.texts.some(text => text.text === 'y = 2x + 1')).toBe(true);
  });

  it('applies the authored major-gridline dash without leaking it to other strokes', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [1, 3, 5, 7] })],
      valAxisGridlineDash: 'dash',
    }), RECT, 1);
    expect(rec.segs.some(segment => segment.dashed)).toBe(true);
    expect(rec.segs.some(segment => !segment.dashed)).toBe(true);
  });

  it('applies minor-gridline paint independently from the major gridlines', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B', 'C', 'D'],
      series: [series({ name: 'S', values: [0, 2, 4, 6] })],
      valMin: 0,
      valMax: 6,
      valAxisMajorUnit: 2,
      valAxisMinorUnit: 1,
      valAxisMinorGridlines: true,
      valAxisMinorGridlineDash: 'dot',
    }), RECT, 1);
    expect(rec.segs.some(segment => segment.dashed)).toBe(true);
    expect(rec.segs.some(segment => !segment.dashed)).toBe(true);
  });

  it('a movingAvg trendline without prstDash draws an additional solid line', () => {
    const noTrend = dashSegRecordingCtx();
    renderChart(noTrend.ctx, lineWithTrend({}), RECT, 1);
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({ trendLines: [{ trendlineType: 'movingAvg', period: 2 }] }), RECT, 1);
    expect(rec.segs.length).toBeGreaterThan(noTrend.segs.length);
    expect(rec.segs.some(s => s.dashed)).toBe(false);
  });

  it.each([
    { trendlineType: 'exp' },
    { trendlineType: 'log' },
    { trendlineType: 'power' },
    { trendlineType: 'poly', order: 2 },
  ])('draws the bounded $trendlineType trendline implementation', trendline => {
    const noTrend = dashSegRecordingCtx();
    renderChart(noTrend.ctx, lineWithTrend({}), RECT, 1);
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({ trendLines: [trendline] }), RECT, 1);
    expect(rec.segs.length).toBeGreaterThan(noTrend.segs.length);
  });

  it('no trendLines field is byte-stable (no dashed segments)', () => {
    const rec = dashSegRecordingCtx();
    renderChart(rec.ctx, lineWithTrend({}), RECT, 1);
    expect(rec.segs.every(s => !s.dashed)).toBe(true);
  });
});

describe('ofPie secondary plots (§21.2.2.126)', () => {
  const ofPieModel = (type: 'pie' | 'bar'): ChartModel => baseModel({
    chartType: 'ofPie',
    categories: ['A', 'B', 'C', 'D', 'E', 'F'],
    series: [series({ values: [40, 30, 20, 10, 5, 2] })],
    ofPie: {
      type,
      splitType: 'pos',
      splitPos: 2,
      secondPieSizePercent: 75,
      gapWidthPercent: 100,
      seriesLines: true,
    },
  });

  it('draws the position split as a primary aggregate plus a secondary pie', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, ofPieModel('pie'), RECT, 1);
    const centers = new Set(rec.arcs.map(arc => `${arc.x.toFixed(2)},${arc.y.toFixed(2)}`));
    expect(centers.size).toBe(2);
    // Four primary points + one aggregate, then the two secondary points.
    expect(rec.arcs).toHaveLength(7);
  });

  it('draws the secondary detail as a stacked bar for bar-of-pie', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, ofPieModel('bar'), RECT, 1);
    const detailBars = rec.rects.filter(rect =>
      ['#70AD47', '#4BACC6'].includes(rect.fs) && rect.w > 5 && rect.h > 0
    );
    expect(detailBars).toHaveLength(2);
  });

  it('uses the Office omission rule without a fixed three-point tail', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      ...ofPieModel('pie'),
      ofPie: {
        ...ofPieModel('pie').ofPie as NonNullable<ChartModel['ofPie']>,
        splitType: 'auto',
        splitTypeAuthored: false,
        splitPos: null,
      },
    }), RECT, 1);
    const arcsPerCenter = new Map<string, number>();
    for (const arc of rec.arcs) {
      const center = `${arc.x.toFixed(2)},${arc.y.toFixed(2)}`;
      arcsPerCenter.set(center, (arcsPerCenter.get(center) ?? 0) + 1);
    }
    // Six points => ceil(6 / 3) = two details. The primary has four source
    // slices plus the aggregate; the secondary has exactly two slices.
    expect([...arcsPerCenter.values()].sort((a, b) => a - b)).toEqual([2, 5]);
  });

  it('does not invent a secondary split for Office-prohibited explicit auto', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      ...ofPieModel('pie'),
      ofPie: {
        ...ofPieModel('pie').ofPie as NonNullable<ChartModel['ofPie']>,
        splitType: 'auto',
        splitTypeAuthored: true,
        splitPos: null,
      },
    }), RECT, 1);
    const centers = new Set(rec.arcs.map(arc => `${arc.x.toFixed(2)},${arc.y.toFixed(2)}`));
    expect(centers.size).toBe(1);
    expect(rec.arcs).toHaveLength(6);
  });

  it.each([
    { splitType: 'pos' as const, splitPos: 6, customSplitIndices: null },
    { splitType: 'cust' as const, splitPos: null, customSplitIndices: [0, 1, 2, 3, 4, 5] },
  ])('keeps an aggregate-only primary plot when $splitType selects every point', split => {
    const rec = recordingCtx();
    const model = ofPieModel('pie');
    renderChart(rec.ctx, {
      ...model,
      ofPie: { ...model.ofPie as NonNullable<ChartModel['ofPie']>, ...split },
    }, RECT, 1);
    const arcsPerCenter = new Map<string, number>();
    for (const arc of rec.arcs) {
      const center = `${arc.x.toFixed(2)},${arc.y.toFixed(2)}`;
      arcsPerCenter.set(center, (arcsPerCenter.get(center) ?? 0) + 1);
    }
    expect([...arcsPerCenter.values()].sort((a, b) => a - b)).toEqual([1, 6]);
  });
});

describe('classic line-chart group decorations', () => {
  const decoratedLine = (): ChartModel => baseModel({
    chartType: 'line',
    categories: ['Day1', 'Day2', 'Day3', 'Day4', 'Day5'],
    valMin: 0,
    valMax: 150,
    valAxisMajorGridlines: false,
    series: [
      series({
        name: 'Open', values: [100, 110, 105, 120, 115],
        lineGroupIndex: 0, showMarker: false, lineColor: '4472C4',
      }),
      series({
        name: 'Close', values: [115, 105, 125, 110, 130],
        lineGroupIndex: 0, showMarker: false, lineColor: 'ED7D31',
      }),
    ],
    lineGroupDecorations: [{
      groupIndex: 0,
      dropLines: { color: '111111', widthEmu: 9525 },
      hiLowLines: { color: '222222', widthEmu: 12700 },
      upDownBars: {
        gapWidthPercent: 150,
        up: { fillColor: 'EEEEEE', lineColor: '333333', lineWidthEmu: 9525 },
        down: { fillColor: '444444', lineColor: '333333', lineWidthEmu: 9525 },
      },
    }],
  });

  it('draws drop lines and high-low lines behind the owning line group', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, decoratedLine(), RECT, 1);
    const vertical = (color: string): Seg[] => rec.segs.filter(segment =>
      segment.ss === color && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) > 1
    );
    // Office vector output has one group-owned envelope per category. It does
    // not paint one coincident drop line for every member series.
    expect(vertical('#111111')).toHaveLength(5);
    expect(vertical('#222222')).toHaveLength(5);
    const firstDecoration = rec.segs.findIndex(segment => segment.ss === '#111111');
    const firstSeries = rec.segs.findIndex(segment => segment.ss === '#4472C4');
    expect(firstDecoration).toBeGreaterThanOrEqual(0);
    expect(firstSeries).toBeGreaterThan(firstDecoration);
  });

  it('places stacked decorations at the plotted cumulative series values', () => {
    const rec = segRecordingCtx();
    const model = decoratedLine();
    model.chartType = 'stackedLine';
    model.series[0].values = [10, 20, 30, 40, 50];
    model.series[1].values = [30, 40, 20, 10, 5];
    model.lineGroupDecorations![0].dropLines = null;
    model.lineGroupDecorations![0].upDownBars = null;
    renderChart(rec.ctx, model, RECT, 1);

    const firstBlue = rec.segs.find(segment => segment.ss === '#4472C4');
    const firstOrange = rec.segs.find(segment => segment.ss === '#ED7D31');
    const firstHiLow = rec.segs.find(segment =>
      segment.ss === '#222222' && Math.abs(segment.x1 - segment.x0) < 0.01
    );
    expect(firstBlue).toBeDefined();
    expect(firstOrange).toBeDefined();
    expect(firstHiLow).toBeDefined();
    const plottedSeriesYs = [firstBlue!.y0, firstOrange!.y0].sort((a, b) => a - b);
    const decorationYs = [firstHiLow!.y0, firstHiLow!.y1].sort((a, b) => a - b);
    expect(decorationYs[0]).toBeCloseTo(plottedSeriesYs[0], 6);
    expect(decorationYs[1]).toBeCloseTo(plottedSeriesYs[1], 6);
  });

  it('draws direct up/down bar paint with the authored gap geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, decoratedLine(), RECT, 1);
    expect(rec.rects.filter(rect => rect.fs === '#EEEEEE')).toHaveLength(3);
    expect(rec.rects.filter(rect => rect.fs === '#444444')).toHaveLength(2);
    const outlines = rec.strokeRects.filter(rect => rect.ss === '#333333');
    expect(outlines).toHaveLength(5);
    expect(new Set(outlines.map(rect => rect.w.toFixed(6))).size).toBe(1);
    const firstSeries = rec.paintEvents.findIndex(event =>
      event.kind === 'stroke' && event.strokeStyle === '#4472C4'
    );
    const firstBar = rec.paintEvents.findIndex(event =>
      event.kind === 'rect' && event.fillStyle === '#EEEEEE'
    );
    expect(firstSeries).toBeGreaterThanOrEqual(0);
    expect(firstBar).toBeGreaterThan(firstSeries);
  });

  it('fills missing line-group decoration paint from linked Chart Style roles', () => {
    const model = decoratedLine();
    model.lineGroupDecorations![0] = {
      groupIndex: 0,
      dropLines: {},
      hiLowLines: {},
      upDownBars: { gapWidthPercent: 150, up: {}, down: {} },
    };
    model.chartStyleRoles = {
      dropLine: { lineColors: ['AA0000'], lineWidthEmu: 19050 },
      hiLoLine: { lineColors: ['00AA00'], lineWidthEmu: 28575 },
      upBar: {
        fillColors: ['AABBCC'], lineColors: ['112233'], lineWidthEmu: 19050,
        lineDash: 'dash', lineCap: 'sq', lineJoin: 'round',
      },
      downBar: {
        fillColors: ['DDEEFF'], lineColors: ['445566'], lineWidthEmu: 28575,
        lineDash: 'dot', lineCap: 'rnd', lineJoin: 'bevel',
      },
    };

    const lines = segRecordingCtx();
    renderChart(lines.ctx, model, RECT, 1);
    expect(lines.segs.filter(segment => segment.ss === '#AA0000')).toHaveLength(5);
    expect(lines.segs.filter(segment => segment.ss === '#00AA00')).toHaveLength(5);

    const bars = recordingCtx();
    renderChart(bars.ctx, model, RECT, 1);
    expect(bars.rects.filter(rect => rect.fs === '#AABBCC')).toHaveLength(3);
    expect(bars.rects.filter(rect => rect.fs === '#DDEEFF')).toHaveLength(2);
    expect(bars.strokeRects.filter(rect => rect.ss === '#112233'
      && rect.dash.length > 0 && rect.cap === 'square' && rect.join === 'round')).toHaveLength(3);
    expect(bars.strokeRects.filter(rect => rect.ss === '#445566'
      && rect.dash.length > 0 && rect.cap === 'round' && rect.join === 'bevel')).toHaveLength(2);
  });

  it('keeps direct decoration paint above linked roles and ignores NoStyle roles', () => {
    const model = decoratedLine();
    model.chartStyleRoles = {
      dropLine: { lineColors: ['AA0000'], lineWidthEmu: 28575 },
      hiLoLine: { lineColors: ['00AA00'], lineNoStyle: true },
      upBar: {
        fillPaints: [{
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
        }],
      },
    };
    const rec = segRecordingCtx();
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.segs.filter(segment => segment.ss === '#111111')).toHaveLength(5);
    expect(rec.segs.some(segment => segment.ss === '#AA0000')).toBe(false);
    expect(rec.segs.some(segment => segment.ss === '#00AA00')).toBe(false);
    const bars = recordingCtx();
    renderChart(bars.ctx, model, RECT, 1);
    expect(bars.gradients).toHaveLength(0);
    expect(bars.rects.filter(rect => rect.fs === '#EEEEEE')).toHaveLength(3);
  });

  it('limits empty up/down-bar automatic paint to the observed classic Style 2 boundary', () => {
    const render = (legacyChartStyle: number | null): Recorded => {
      const rec = recordingCtx();
      const model = decoratedLine();
      model.legacyChartStyle = legacyChartStyle;
      model.lineGroupDecorations![0].upDownBars = {
        gapWidthPercent: 150, up: {}, down: {},
      };
      renderChart(rec.ctx, model, RECT, 1);
      return rec;
    };
    const styleTwo = render(2);
    expect(styleTwo.rects.filter(rect => rect.fs === '#FFFFFF')).toHaveLength(3);
    expect(styleTwo.rects.filter(rect => rect.fs === '#000000')).toHaveLength(2);

    const unresolvedStyle = render(null);
    expect(unresolvedStyle.rects.filter(rect =>
      rect.fs === '#FFFFFF' || rect.fs === '#000000'
    )).toHaveLength(0);
    expect(unresolvedStyle.strokeRects).toHaveLength(0);
  });

  it('keeps decorations scoped to their owning line group', () => {
    const rec = segRecordingCtx();
    const model = decoratedLine();
    model.series.push(series({
      name: 'Other group', values: [50, 55, 60, 65, 70],
      lineGroupIndex: 1, showMarker: false, lineColor: '70AD47',
    }));
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.segs.filter(segment =>
      segment.ss === '#111111' && Math.abs(segment.x1 - segment.x0) < 0.01
    )).toHaveLength(5);
  });

  it('uses one interior crossing for the axis, labels, and drop-line envelopes', () => {
    const rec = segRecordingCtx();
    const model = decoratedLine();
    model.catAxisCrossesAt = 75;
    model.catAxisLineColor = 'ABCDEF';
    model.catAxisMajorTickMark = 'none';
    model.series[0].values = [100, 110, 105, 120, 115];
    model.series[1].values = [115, 105, 125, 110, 130];

    renderChart(rec.ctx, model, RECT, 1);

    const axis = rec.segs.find(segment =>
      segment.ss === '#ABCDEF'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100
    );
    expect(axis).toBeDefined();
    const dropLines = rec.segs.filter(segment =>
      segment.ss === '#111111'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) > 1
    );
    expect(dropLines).toHaveLength(5);
    for (const line of dropLines) {
      expect(Math.max(line.y0, line.y1)).toBeCloseTo(axis!.y0, 6);
    }
    const categoryLabels = rec.texts.filter(text => text.text.startsWith('Day'));
    expect(categoryLabels).toHaveLength(5);
    expect(categoryLabels.every(text => text.y > axis!.y0)).toBe(true);
  });

  it('keeps low category labels at the plot edge when the axis crosses inside', () => {
    const render = (position: 'nextTo' | 'low'): { axisY: number; labelY: number } => {
      const rec = segRecordingCtx();
      const model = decoratedLine();
      model.catAxisCrossesAt = 75;
      model.catAxisLineColor = 'ABCDEF';
      model.catAxisTickLabelPos = position;
      model.catAxisMajorTickMark = 'none';
      renderChart(rec.ctx, model, RECT, 1);
      const axis = rec.segs.find(segment =>
        segment.ss === '#ABCDEF' && Math.abs(segment.y1 - segment.y0) < 0.01
      );
      const label = rec.texts.find(text => text.text === 'Day1');
      expect(axis).toBeDefined();
      expect(label).toBeDefined();
      return { axisY: axis!.y0, labelY: label!.y };
    };

    const nextTo = render('nextTo');
    const low = render('low');
    expect(nextTo.labelY).toBeGreaterThan(nextTo.axisY);
    expect(low.labelY).toBeGreaterThan(nextTo.labelY);
  });

  it('maps min and max crossings through the authored value-axis orientation', () => {
    const axisY = (
      crossing: 'min' | 'max',
      orientation: 'minMax' | 'maxMin',
    ): number => {
      const rec = segRecordingCtx();
      const model = decoratedLine();
      model.valMin = -10;
      model.valMax = 10;
      model.valAxisOrientation = orientation;
      model.catAxisCrosses = crossing;
      model.catAxisLineColor = 'ABCDEF';
      model.catAxisMajorTickMark = 'none';
      renderChart(rec.ctx, model, RECT, 1);
      const axis = rec.segs.find(segment =>
        segment.ss === '#ABCDEF'
        && Math.abs(segment.y1 - segment.y0) < 0.01
        && Math.abs(segment.x1 - segment.x0) > 100
      );
      expect(axis).toBeDefined();
      return axis!.y0;
    };

    const minimum = axisY('min', 'minMax');
    const maximum = axisY('max', 'minMax');
    expect(minimum).toBeGreaterThan(maximum);
    expect(axisY('min', 'maxMin')).toBeCloseTo(maximum, 6);
    expect(axisY('max', 'maxMin')).toBeCloseTo(minimum, 6);
  });

  it('uses the paired secondary crossing for a secondary line group', () => {
    const dropLength = (crossesAt: number): number => {
      const rec = segRecordingCtx();
      const model = baseModel({
        chartType: 'line', categories: ['A'], valMin: 0, valMax: 10,
        valAxisMajorGridlines: false,
        series: [series({
          values: [150], useSecondaryAxis: true, lineGroupIndex: 1,
          showMarker: false, lineColor: '4472C4',
        })],
        secondaryValAxis: {
          min: 0, max: 200, title: null, hidden: true,
          lineHidden: true, majorTickMark: 'none',
        },
        secondaryCatAxis: {
          min: null, max: null, title: null, hidden: true,
          lineHidden: true, majorTickMark: 'none', crossesAt,
        },
        lineGroupDecorations: [{
          groupIndex: 1,
          dropLines: { color: '123456', widthEmu: 9525 },
        }],
      });
      renderChart(rec.ctx, model, RECT, 1);
      const drop = rec.segs.find(segment =>
        segment.ss === '#123456' && Math.abs(segment.x1 - segment.x0) < 0.01
      );
      expect(drop).toBeDefined();
      return Math.abs(drop!.y1 - drop!.y0);
    };

    expect(dropLength(100)).toBeLessThan(dropLength(0));
  });
});

describe('classic area-chart group drop lines', () => {
  it('draws one authored envelope line per category for a multi-series group', () => {
    const rec = segRecordingCtx();
    const model = baseModel({
      chartType: 'area',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 30,
      valAxisMajorGridlines: false,
      series: [
        series({ name: 'First', values: [10, 20, 15], lineColor: '4472C4' }),
        series({ name: 'Second', values: [15, 5, 25], lineColor: 'ED7D31' }),
      ],
    });
    const extended = model as ChartModel & {
      areaGroupDecorations?: Array<{
        groupIndex: number;
        dropLines?: { color?: string | null; widthEmu?: number | null } | null;
      }>;
    };
    extended.areaGroupDecorations = [{
      groupIndex: 0,
      dropLines: { color: '123456', widthEmu: 12700 },
    }];
    for (const item of extended.series) {
      (item as ChartSeries & { areaGroupIndex?: number | null }).areaGroupIndex = 0;
    }

    renderChart(rec.ctx, extended, RECT, 1);

    const dropLines = rec.segs.filter(segment =>
      segment.ss === '#123456'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) > 1
    );
    expect(dropLines).toHaveLength(3);
    expect(new Set(dropLines.map(line => line.x0.toFixed(6))).size).toBe(3);
  });

  it('uses the cumulative plotted value for percent and stacked area groups', () => {
    const rec = segRecordingCtx();
    const model = baseModel({
      chartType: 'stackedArea',
      categories: ['A'],
      valMin: 0,
      valMax: 30,
      valAxisMajorGridlines: true,
      valAxisMajorUnit: 30,
      series: [
        series({ name: 'First', values: [10], lineColor: '4472C4' }),
        series({ name: 'Second', values: [20], lineColor: 'ED7D31' }),
      ],
    });
    const extended = model as ChartModel & {
      areaGroupDecorations?: Array<{
        groupIndex: number;
        dropLines?: { color?: string | null; widthEmu?: number | null } | null;
      }>;
    };
    extended.areaGroupDecorations = [{ groupIndex: 0, dropLines: { color: '123456' } }];
    for (const item of extended.series) {
      (item as ChartSeries & { areaGroupIndex?: number | null }).areaGroupIndex = 0;
    }

    renderChart(rec.ctx, extended, RECT, 1);

    const dropLines = rec.segs.filter(segment => segment.ss === '#123456');
    expect(dropLines).toHaveLength(1);
    const topGridlineY = Math.min(
      ...rec.segs
        .filter(segment => Math.abs(segment.y1 - segment.y0) < 0.01
          && Math.abs(segment.x1 - segment.x0) > 100)
        .map(segment => segment.y0),
    );
    expect(Math.min(dropLines[0].y0, dropLines[0].y1)).toBeCloseTo(topGridlineY, 6);
  });

  it('starts drop lines at an explicitly crossed category axis', () => {
    const render = (crossesAt: number): Seg[] => {
      const rec = segRecordingCtx();
      const model = baseModel({
        chartType: 'area', categories: ['A', 'B'], valMin: 0, valMax: 30,
        catAxisCrossesAt: crossesAt, valAxisMajorGridlines: false,
        series: [
          series({ name: 'First', values: [5, 14] }),
          series({ name: 'Second', values: [18, 22] }),
        ],
      });
      const extended = model as ChartModel & {
        areaGroupDecorations?: Array<{
          groupIndex: number;
          dropLines?: { color?: string | null } | null;
        }>;
      };
      extended.areaGroupDecorations = [{ groupIndex: 0, dropLines: { color: '123456' } }];
      for (const item of extended.series) {
        (item as ChartSeries & { areaGroupIndex?: number | null }).areaGroupIndex = 0;
      }
      renderChart(rec.ctx, extended, RECT, 1);
      return rec.segs.filter(segment => segment.ss === '#123456');
    };

    const atZero = render(0);
    const atTen = render(10);
    expect(atZero).toHaveLength(2);
    expect(atTen).toHaveLength(2);
    expect(Math.abs(atTen[0].y1 - atTen[0].y0))
      .toBeLessThan(Math.abs(atZero[0].y1 - atZero[0].y0));
  });

  it('moves the visible axis, ticks, and next-to labels to the same interior crossing', () => {
    const rec = segRecordingCtx();
    const model = baseModel({
      chartType: 'area', categories: ['A', 'B'], valMin: 0, valMax: 30,
      catAxisCrossesAt: 10, catAxisLineColor: 'ABCDEF',
      catAxisMajorTickMark: 'out', valAxisMajorGridlines: false,
      series: [series({ values: [18, 22] })],
    });
    renderChart(rec.ctx, model, RECT, 1);

    const axis = rec.segs.find(segment =>
      segment.ss === '#ABCDEF'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) > 100
    );
    expect(axis).toBeDefined();
    const ticks = rec.segs.filter(segment =>
      segment.ss === '#ABCDEF'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) > 0
    );
    expect(ticks.length).toBeGreaterThan(0);
    expect(ticks.every(tick => Math.min(tick.y0, tick.y1) >= axis!.y0 - 0.01)).toBe(true);
    const labels = rec.texts.filter(text => text.text === 'A' || text.text === 'B');
    expect(labels).toHaveLength(2);
    expect(labels.every(text => text.y > axis!.y0)).toBe(true);
  });
});

describe('classic bar-chart group series lines', () => {
  const decorate = (model: ChartModel): ChartModel => {
    model.barGroupDecorations = [{
      groupIndex: 0,
      seriesLines: [{ color: '234567', widthEmu: 19050 }],
    }];
    for (const item of model.series) {
      item.barGroupIndex = 0;
      item.barGroupGrouping = 'stacked';
    }
    return model;
  };

  it('fills a missing series-line paint from the linked Chart Style role', () => {
    const rec = segRecordingCtx();
    const model = decorate(baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 50,
      valAxisMajorGridlines: false,
      series: [series({ values: [10, 20, 15], barGroupDirection: 'col' })],
    }));
    model.barGroupDecorations![0].seriesLines = [{}];
    model.chartStyleRoles = {
      seriesLine: { lineColors: ['765432'], lineWidthEmu: 19050 },
    };

    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.segs.filter(segment => segment.ss === '#765432')).toHaveLength(2);
  });

  it('joins the facing column edges for every adjacent point in each series', () => {
    const rec = segRecordingCtx();
    const model = decorate(baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 50,
      valAxisMajorGridlines: false,
      series: [
        series({ name: 'First', values: [10, 20, 15], barGroupDirection: 'col' }),
        series({ name: 'Second', values: [5, 10, 20], barGroupDirection: 'col' }),
      ],
    }));

    renderChart(rec.ctx, model, RECT, 1);

    const lines = rec.segs.filter(segment => segment.ss === '#234567');
    expect(lines).toHaveLength(4);
    const centers = rec.texts
      .filter(text => ['A', 'B', 'C'].includes(text.text))
      .map(text => text.x)
      .sort((left, right) => left - right);
    expect(centers).toHaveLength(3);
    for (const line of lines) {
      const left = Math.min(line.x0, line.x1);
      const right = Math.max(line.x0, line.x1);
      expect(centers.some((center, index) => index + 1 < centers.length
        && left > center && right < centers[index + 1])).toBe(true);
    }
  });

  it('joins facing horizontal-bar edges and keeps negative value endpoints', () => {
    const rec = segRecordingCtx();
    const model = decorate(baseModel({
      chartType: 'stackedBarH',
      categories: ['A', 'B', 'C'],
      valMin: -40,
      valMax: 0,
      valAxisMajorGridlines: false,
      series: [series({
        name: 'Negative', values: [-10, -25, -15], barGroupDirection: 'bar',
      })],
    }));

    renderChart(rec.ctx, model, RECT, 1);

    const lines = rec.segs.filter(segment => segment.ss === '#234567');
    expect(lines).toHaveLength(2);
    const centers = rec.texts
      .filter(text => ['A', 'B', 'C'].includes(text.text))
      .map(text => text.y)
      .sort((top, bottom) => top - bottom);
    expect(centers).toHaveLength(3);
    for (const line of lines) {
      const top = Math.min(line.y0, line.y1);
      const bottom = Math.max(line.y0, line.y1);
      expect(centers.some((center, index) => index + 1 < centers.length
        && top > center && bottom < centers[index + 1])).toBe(true);
      expect(line.x0).toBeLessThan(RECT.x + RECT.w);
      expect(line.x1).toBeLessThan(RECT.x + RECT.w);
    }
  });

  it('breaks a series line across a missing data point', () => {
    const rec = segRecordingCtx();
    const model = decorate(baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B', 'C'],
      valMin: 0,
      valMax: 30,
      valAxisMajorGridlines: false,
      series: [series({
        name: 'Sparse', values: [10, null, 20], barGroupDirection: 'col',
      })],
    }));

    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.segs.filter(segment => segment.ss === '#234567')).toHaveLength(0);
  });

  it('keeps multiple authored series-line styles unrendered until association is verified', () => {
    const rec = segRecordingCtx();
    const model = decorate(baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B'],
      valMin: 0,
      valMax: 30,
      valAxisMajorGridlines: false,
      series: [series({ name: 'Series', values: [10, 20], barGroupDirection: 'col' })],
    }));
    model.barGroupDecorations![0].seriesLines!.push({ color: 'FF0000' });

    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.segs.filter(segment =>
      segment.ss === '#234567' || segment.ss === '#FF0000'
    )).toHaveLength(0);
  });
});

describe('CH13 — stock chart (high/low/close)', () => {
  // High/Low/Close over three dates. Value axis 0..70 so the plot geometry is
  // easy to reason about.
  const stockModel = (over: Partial<ChartModel> = {}): ChartModel => baseModel({
    chartType: 'stock',
    categories: ['1/5/2002', '1/6/2002', '1/7/2002'],
    valMin: 0,
    valMax: 70,
    stockHiLowLines: true,
    stockHiLowLineColor: '595959',
    stockAutomaticStyle: {
      lineColor: '000000', lineWidthEmu: 12700,
      upFillColor: 'F9F9F9', downFillColor: '3F3F3F',
    },
    series: [
      series({ name: 'High', values: [55, 57, 57] }),
      series({ name: 'Low', values: [11, 12, 13] }),
      series({ name: 'Close', values: [32, 35, 34] }),
    ],
    ...over,
  });

  /** Near-vertical segments in the hi-lo line color that span a large Y range —
   *  these are the per-category low↔high lines. */
  function hiLoLines(segs: Seg[]): Seg[] {
    return segs.filter(
      s => Math.abs(s.x1 - s.x0) < 0.5 && Math.abs(s.y1 - s.y0) > 20 && s.ss === '#595959',
    );
  }

  it('draws one vertical low↔high line per category, spanning the correct value range', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel(), RECT, 1);
    const lines = hiLoLines(rec.segs);
    expect(lines.length).toBe(3);

    // The tallest line (category 2: 12..57 = 45 units) is taller than the
    // shortest of the three, and every line's two endpoints map High above Low
    // (smaller Y = higher value in canvas coords).
    for (const l of lines) {
      const top = Math.min(l.y0, l.y1);
      const bot = Math.max(l.y0, l.y1);
      expect(bot).toBeGreaterThan(top);
    }
    // Category ordering left→right: the three lines have increasing X.
    const xs = lines.map(l => l.x0).sort((a, b) => a - b);
    expect(xs[0]).toBeLessThan(xs[1]);
    expect(xs[1]).toBeLessThan(xs[2]);

    // The hi-lo span is proportional to (high - low): category 1 (55-11=44) is
    // shorter than category 2 (57-12=45) by roughly the same pixel ratio.
    const span = (l: Seg): number => Math.abs(l.y1 - l.y0);
    const byX = [...lines].sort((a, b) => a.x0 - b.x0);
    // 44 vs 45 vs 44 units — cat2 is the tallest.
    expect(span(byX[1])).toBeGreaterThanOrEqual(span(byX[0]));
  });

  it('renders the title and a series-driven legend (High / Low / Close)', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({ title: 'Stock', showLegend: true, legendPos: 'b' }), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    expect(labels).toContain('Stock');
    expect(labels).toContain('High');
    expect(labels).toContain('Low');
    expect(labels).toContain('Close');
  });

  it('keeps cached date labels when dateAx leaves the automatic major interval omitted', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({
      categories: ['43851', '43852', '43853'],
      catAxisIsDate: true,
      catAxisBaseTimeUnit: 'days',
      catAxisFormatCode: 'mm/dd/yyyy',
      catAxisMajorUnit: null,
      catAxisTickLabelPos: 'nextTo',
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining([
      '01/21/2020',
      '01/22/2020',
      '01/23/2020',
    ]));
  });

  it('honors an explicit hi-lo line color', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({ stockHiLowLineColor: 'FF0000' }), RECT, 1);
    const red = rec.segs.filter(
      s => Math.abs(s.x1 - s.x0) < 0.5 && Math.abs(s.y1 - s.y0) > 20 && s.ss === '#FF0000',
    );
    expect(red.length).toBe(3);
  });

  it('honors complete high-low line paint, noFill, and linked Chart Style fallback', () => {
    const direct = segRecordingCtx();
    renderChart(direct.ctx, stockModel({
      stockHiLowLineStyle: {
        color: 'AA0000', widthEmu: 25400, dash: 'dot', cap: 'rnd', join: 'bevel',
      },
    }), RECT, 1);
    const directLines = direct.segs.filter(segment => segment.ss === '#AA0000');
    expect(directLines).toHaveLength(3);
    expect(directLines.every(segment => segment.lw === 2 && segment.dash.length > 0
      && segment.cap === 'round' && segment.join === 'bevel')).toBe(true);

    const hidden = segRecordingCtx();
    renderChart(hidden.ctx, stockModel({
      stockHiLowLineStyle: { hidden: true },
      chartStyleRoles: { hiLoLine: { lineColors: ['00AA00'], lineWidthEmu: 25400 } },
    }), RECT, 1);
    expect(hidden.segs.some(segment => segment.ss === '#00AA00')).toBe(false);

    const linked = segRecordingCtx();
    renderChart(linked.ctx, stockModel({
      stockHiLowLineStyle: {},
      chartStyleRoles: {
        hiLoLine: { lineColors: ['00AA00'], lineWidthEmu: 19050, lineDash: 'dash' },
      },
    }), RECT, 1);
    const linkedLines = linked.segs.filter(segment => segment.ss === '#00AA00');
    expect(linkedLines).toHaveLength(3);
    expect(linkedLines.every(segment => segment.lw === 1.5 && segment.dash.length > 0)).toBe(true);
  });

  it('draws one styled stock drop-line envelope per category', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({
      stockDropLines: {
        color: '123456', widthEmu: 12700, dash: 'dashDot', cap: 'sq', join: 'round',
      },
    }), RECT, 1);

    const dropLines = rec.segs.filter(segment => segment.ss === '#123456');
    expect(dropLines).toHaveLength(3);
    expect(dropLines.every(segment => Math.abs(segment.x1 - segment.x0) < 0.01)).toBe(true);
    expect(dropLines.every(segment => Math.abs(segment.y1 - segment.y0) > 20)).toBe(true);
    expect(dropLines.every(segment => segment.lw === 1 && segment.dash.length > 0
      && segment.cap === 'square' && segment.join === 'round')).toBe(true);
  });

  it('keeps stock drop-line noFill hidden and resolves omitted paint from Chart Style', () => {
    const hidden = segRecordingCtx();
    renderChart(hidden.ctx, stockModel({
      stockDropLines: { hidden: true },
      chartStyleRoles: { dropLine: { lineColors: ['AABBCC'], lineWidthEmu: 25400 } },
    }), RECT, 1);
    expect(hidden.segs.some(segment => segment.ss === '#AABBCC')).toBe(false);

    const linked = segRecordingCtx();
    renderChart(linked.ctx, stockModel({
      stockDropLines: {},
      chartStyleRoles: {
        dropLine: {
          lineColors: ['AABBCC'], lineWidthEmu: 25400, lineDash: 'dash',
          lineCap: 'rnd', lineJoin: 'bevel',
        },
      },
    }), RECT, 1);
    const dropLines = linked.segs.filter(segment => segment.ss === '#AABBCC');
    expect(dropLines).toHaveLength(3);
    expect(dropLines.every(segment => segment.lw === 2 && segment.dash.length > 0
      && segment.cap === 'round' && segment.join === 'bevel')).toBe(true);
  });

  it('uses the bounded automatic line recipe for an empty stock drop-line element', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({
      stockHiLowLines: false,
      stockDropLines: {},
      stockAutomaticStyle: {
        lineColor: '240000', lineWidthEmu: 12700,
        upFillColor: 'F9F9F9', downFillColor: '493F3F',
      },
    }), RECT, 1);
    const lines = rec.segs.filter(segment => segment.ss === '#240000');
    expect(lines).toHaveLength(3);
    expect(lines.every(segment => segment.lw === 1)).toBe(true);

    const unresolved = segRecordingCtx();
    renderChart(unresolved.ctx, stockModel({
      stockHiLowLines: false,
      stockDropLines: {},
      stockAutomaticStyle: null,
    }), RECT, 1);
    expect(unresolved.segs.some(segment => segment.ss === '#000000')).toBe(false);

    for (const stockDropLines of [
      { paintAuthored: true, widthEmu: 12700 },
      {},
    ]) {
      const authoredUnresolved = segRecordingCtx();
      renderChart(authoredUnresolved.ctx, stockModel({
        stockHiLowLines: false,
        stockDropLines,
        chartStyleRoles: stockDropLines.paintAuthored === true ? undefined : {
          dropLine: { linePaintAuthored: true, lineWidthEmu: 12700 },
        },
        stockAutomaticStyle: {
          lineColor: '240000', lineWidthEmu: 12700,
          upFillColor: 'F9F9F9', downFillColor: '493F3F',
        },
      }), RECT, 1);
      expect(authoredUnresolved.segs.some(segment => segment.ss === '#240000')).toBe(false);
    }

    const linkedGeometryOnly = segRecordingCtx();
    renderChart(linkedGeometryOnly.ctx, stockModel({
      stockHiLowLines: false,
      stockDropLines: {},
      chartStyleRoles: {
        dropLine: { lineWidthEmu: 25400, lineDash: 'dash' },
      },
      stockAutomaticStyle: {
        lineColor: '240000', lineWidthEmu: 12700,
        upFillColor: 'F9F9F9', downFillColor: '493F3F',
      },
    }), RECT, 1);
    const geometryLines = linkedGeometryOnly.segs.filter(segment => segment.ss === '#240000');
    expect(geometryLines).toHaveLength(3);
    expect(geometryLines.every(segment => segment.lw === 2 && segment.dash.length > 0)).toBe(true);
  });

  it('uses parser-resolved automatic stock line paint only when the element exists', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({
      stockHiLowLineColor: null,
      stockHiLowLineStyle: {},
      stockAutomaticStyle: {
        lineColor: '240000', lineWidthEmu: 12700,
        upFillColor: 'F9F9F9', downFillColor: '493F3F',
      },
    }), RECT, 1);
    const automatic = rec.segs.filter(segment => segment.ss === '#240000');
    expect(automatic).toHaveLength(3);
    expect(automatic.every(segment => segment.lw === 1)).toBe(true);

    const absent = segRecordingCtx();
    renderChart(absent.ctx, stockModel({
      stockHiLowLines: undefined,
      stockHiLowLineColor: null,
      stockHiLowLineStyle: null,
      stockAutomaticStyle: {
        lineColor: '240000', lineWidthEmu: 12700,
        upFillColor: 'F9F9F9', downFillColor: '493F3F',
      },
    }), RECT, 1);
    expect(absent.segs.some(segment => segment.ss === '#240000')).toBe(false);
  });

  it('maps secondary stock series and its right value axis through one scale', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, stockModel({
      valMin: 0,
      valMax: 10,
      secondaryValAxis: {
        min: 0, max: 100, majorUnit: 50,
        title: null, hidden: false, lineHidden: false,
        lineColor: '00AA00', lineWidthEmu: 12700,
        majorTickMark: 'cross', tickLabelPos: 'nextTo',
      },
      series: [
        series({ name: 'High', values: [55, 57, 57], useSecondaryAxis: true }),
        series({ name: 'Low', values: [11, 12, 13], useSecondaryAxis: true }),
        series({ name: 'Close', values: [32, 35, 34], useSecondaryAxis: true }),
      ],
    }), RECT, 1);

    expect(hiLoLines(rec.segs)).toHaveLength(3);
    expect(hiLoLines(rec.segs).every(line =>
      line.y0 >= RECT.y && line.y0 <= RECT.y + RECT.h
      && line.y1 >= RECT.y && line.y1 <= RECT.y + RECT.h
    )).toBe(true);
    expect(rec.segs.some(line => line.ss === '#00AA00'
      && Math.abs(line.x1 - line.x0) < 0.01
      && Math.abs(line.y1 - line.y0) > 100)).toBe(true);
    expect(rec.texts.map(text => text.text)).toContain('100');
  });

  it('honors a stock-series point marker override without changing sibling ticks', () => {
    const rec = recordingCtx();
    const model = stockModel();
    model.series[2] = series({
      name: 'Close', values: [32, 35, 34],
      dataPointOverrides: [{
        idx: 1, markerSymbol: 'circle', markerSize: 7,
        markerLine: 'AA0000', markerLineWidthEmu: 25400,
        markerFillPaint: {
          fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [
            { position: 0, color: '112233' },
            { position: 1, color: 'DDEEFF' },
          ],
        },
      }],
    });
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.gradients).toHaveLength(1);
    expect(rec.paintEvents.some(event =>
      event.kind === 'stroke' && event.strokeStyle === '#AA0000'
    )).toBe(true);
  });

  it('renders a stock-series picture marker override through the shared image path', () => {
    const rec = recordingCtx();
    const bitmap = { width: 24, height: 24 } as unknown as CanvasImageSource;
    const model = stockModel();
    model.series[2] = series({
      name: 'Close', values: [32, 35, 34],
      dataPointOverrides: [{
        idx: 1, markerSymbol: 'picture', markerSize: 7,
        markerFillPaintAuthored: true,
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'word/media/marker.png', mimeType: 'image/png',
        },
      }],
    });
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(rec.drawImages).toHaveLength(1);
    expect(rec.drawImages[0][0]).toBe(bitmap);
  });

  it('renders a point-only picture marker on the stock High series', () => {
    const rec = recordingCtx();
    const bitmap = { width: 24, height: 24 } as unknown as CanvasImageSource;
    const model = stockModel();
    model.series[0] = series({
      name: 'High', values: [55, 57, 57],
      dataPointOverrides: [{
        idx: 1, markerSymbol: 'picture', markerSize: 7,
        markerFillPaintAuthored: true,
        markerFillPaint: {
          fillType: 'image', stretch: true, imagePath: 'word/media/high-marker.png', mimeType: 'image/png',
        },
      }],
    });
    renderChartCore(rec.ctx, model, RECT, 1, 0, testThreeD, undefined, () => bitmap);
    expect(rec.drawImages).toHaveLength(1);
  });

  it('paints stock-series defaults and per-point rich callout labels', () => {
    const rec = recordingCtx();
    const model = stockModel();
    model.series[2] = series({
      name: 'Close',
      values: [32, 35, 34],
      seriesDataLabels: {
        showVal: true,
        showCatName: true,
        showSerName: false,
        showPercent: false,
        fontColor: '008000',
        separator: ' | ',
      },
      dataLabelOverrides: [{
        idx: 1,
        text: 'Close custom',
        richRuns: [{ text: 'Close custom', italic: true }],
        labelBox: { fill: 'FFF2CC', fillPaintAuthored: true },
      }],
    });

    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['1/5/2002 | 32', 'Close custom', '1/7/2002 | 34']),
    );
    expect(rec.rects.some(rect => rect.fs === '#FFF2CC')).toBe(true);
  });

  it('lets a point-level delete=false restore one stock label from a deleted collection', () => {
    const rec = recordingCtx();
    const model = stockModel();
    model.series[2] = series({
      name: 'Close',
      values: [32, 35, 34],
      seriesDataLabels: {
        deleted: true,
        showVal: true,
        showCatName: false,
        showSerName: false,
        showPercent: false,
      },
      dataLabelOverrides: [{ idx: 1, text: '', deleted: false, showVal: true }],
    });

    renderChart(rec.ctx, model, RECT, 1);

    const texts = rec.texts.map(text => text.text);
    expect(texts).toContain('35');
    expect(texts).not.toContain('32');
    expect(texts).not.toContain('34');
  });

  it('draws authored stock-chart minor ticks', () => {
    const count = (minorTick: ChartModel['valAxisMinorTickMark']): number => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, stockModel({
        valMin: 3, valMax: 23, valAxisMajorUnit: 10, valAxisMinorUnit: 4,
        valAxisMajorGridlines: false, valAxisMinorGridlines: false,
        valAxisMajorTickMark: 'none', valAxisMinorTickMark: minorTick,
      }), RECT, 1);
      return rec.segs.filter(segment =>
        Math.abs(segment.y1 - segment.y0) < 0.01
        && Math.abs(segment.x1 - segment.x0) > 0
        && Math.abs(segment.x1 - segment.x0) <= 12
      ).length;
    };
    expect(count('cross') - count('none')).toBe(4);
  });

  it('draws stock-series error bars and includes their endpoints in auto scaling', () => {
    const rec = recordingCtx();
    const model = stockModel({
      valMin: null,
      valMax: null,
      series: [
        series({
          name: 'High', values: [55, 57, 57],
          errBars: [{
            dir: 'y', barType: 'plus', plus: [25, 25, 25], minus: [],
            noEndCap: true, color: 'FF00FF',
          }],
        }),
        series({ name: 'Low', values: [11, 12, 13] }),
        series({ name: 'Close', values: [32, 35, 34] }),
      ],
    });
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.paintEvents).toContainEqual({ kind: 'stroke', strokeStyle: '#FF00FF' });
    expect(rec.texts.map(text => text.text)).toContain('90');
  });

  it('draws styled open-close up/down bars with the authored gap width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: {
        gapWidthPercent: 100,
        up: {
          fillColor: '00AA00', lineColor: '006600', lineWidthEmu: 12700,
          lineDash: 'dash', lineCap: 'sq', lineJoin: 'round',
        },
        down: {
          fillColor: 'CC0000', lineColor: '660000', lineWidthEmu: 25400,
          lineDash: 'dot', lineCap: 'rnd', lineJoin: 'bevel',
        },
      },
      series: [
        series({ name: 'Open', values: [20, 45, 25] }),
        series({ name: 'High', values: [55, 57, 57] }),
        series({ name: 'Low', values: [11, 12, 13] }),
        series({ name: 'Close', values: [40, 30, 25] }),
      ],
    }), RECT, 1);

    const bars = rec.rects.filter(rect => rect.fs === '#00AA00' || rect.fs === '#CC0000');
    expect(bars).toHaveLength(2); // the equal open/close point has zero area
    expect(bars.map(rect => rect.fs)).toEqual(['#00AA00', '#CC0000']);
    expect(bars[0].w).toBeCloseTo(bars[1].w, 6);
    expect(bars[0].w).toBeGreaterThan(30);
    expect(rec.strokeRects.some(rect => rect.ss === '#006600' && rect.lw === 1
      && rect.dash.length > 0 && rect.cap === 'square' && rect.join === 'round')).toBe(true);
    expect(rec.strokeRects.some(rect => rect.ss === '#660000' && rect.lw === 2
      && rect.dash.length > 0 && rect.cap === 'round' && rect.join === 'bevel')).toBe(true);
  });

  it('applies bounded automatic stock bar paint after direct and linked paint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: {
        gapWidthPercent: 100,
        up: {},
        down: { fillColor: 'CC0000', lineHidden: true },
      },
      stockAutomaticStyle: {
        lineColor: '240000', lineWidthEmu: 12700,
        upFillColor: 'F9F9F9', downFillColor: '493F3F',
      },
      series: [
        series({ name: 'Open', values: [20, 45, 25] }),
        series({ name: 'High', values: [55, 57, 57] }),
        series({ name: 'Low', values: [11, 12, 13] }),
        series({ name: 'Close', values: [40, 30, 25] }),
      ],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#F9F9F9')).toHaveLength(1);
    expect(rec.rects.filter(rect => rect.fs === '#CC0000')).toHaveLength(1);
    expect(rec.strokeRects.filter(rect => rect.ss === '#240000' && rect.lw === 1)).toHaveLength(1);

    const unresolved = recordingCtx();
    renderChart(unresolved.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: { gapWidthPercent: 100, up: {}, down: {} },
      stockAutomaticStyle: null,
      series: [
        series({ name: 'Open', values: [20, 45] }),
        series({ name: 'High', values: [55, 57] }),
        series({ name: 'Low', values: [11, 12] }),
        series({ name: 'Close', values: [40, 30] }),
      ],
    }), RECT, 1);
    expect(unresolved.rects.filter(rect => rect.fs === '#FFFFFF' || rect.fs === '#000000'))
      .toHaveLength(0);
    expect(unresolved.strokeRects.filter(rect => rect.ss === '#000000')).toHaveLength(0);

    const authoredUnresolved = recordingCtx();
    renderChart(authoredUnresolved.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: {
        gapWidthPercent: 100,
        up: {
          fill: { fillType: 'gradient', gradType: 'linear', angle: 0, stops: [] },
        },
        down: {},
      },
      series: [
        series({ name: 'Open', values: [20] }),
        series({ name: 'High', values: [55] }),
        series({ name: 'Low', values: [11] }),
        series({ name: 'Close', values: [40] }),
      ],
    }), RECT, 1);
    expect(authoredUnresolved.rects.filter(rect => rect.fs === '#F9F9F9')).toHaveLength(0);

    const authoredUnresolvedProvenance = recordingCtx();
    renderChart(authoredUnresolvedProvenance.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: {
        gapWidthPercent: 100,
        up: { fillPaintAuthored: true, linePaintAuthored: true },
        down: { fillHidden: true, lineHidden: true },
      },
      series: [
        series({ name: 'Open', values: [20] }),
        series({ name: 'High', values: [55] }),
        series({ name: 'Low', values: [11] }),
        series({ name: 'Close', values: [40] }),
      ],
    }), RECT, 1);
    expect(authoredUnresolvedProvenance.rects.filter(rect => rect.fs === '#F9F9F9'))
      .toHaveLength(0);
    expect(authoredUnresolvedProvenance.strokeRects.filter(rect => rect.ss === '#000000'))
      .toHaveLength(0);
  });

  it('keeps automatic up/down bars correct across zero and missing points', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      valMin: -20,
      valMax: 20,
      stockUpDownBars: true,
      stockUpDownBarStyle: { gapWidthPercent: 100, up: {}, down: {} },
      series: [
        series({ name: 'Open', values: [-10, 10, null] }),
        series({ name: 'High', values: [15, 15, 15] }),
        series({ name: 'Low', values: [-15, -15, -15] }),
        series({ name: 'Close', values: [10, -10, 5] }),
      ],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#F9F9F9')).toHaveLength(1);
    expect(rec.rects.filter(rect => rect.fs === '#3F3F3F')).toHaveLength(1);
    expect(rec.strokeRects.filter(rect => rect.ss === '#000000')).toHaveLength(2);
  });

  it('uses the shared structured-fill renderer for stock up/down bars', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      stockUpDownBars: true,
      stockUpDownBarStyle: {
        gapWidthPercent: 100,
        up: {
          fill: {
            fillType: 'gradient', gradType: 'linear', angle: 90,
            stops: [
              { position: 0, color: '112233' },
              { position: 1, color: 'DDEEFF' },
            ],
          },
        },
        down: { fillColor: 'CC0000' },
      },
      series: [
        series({ name: 'Open', values: [20, 45, 25] }),
        series({ name: 'High', values: [55, 57, 57] }),
        series({ name: 'Low', values: [11, 12, 13] }),
        series({ name: 'Close', values: [40, 30, 25] }),
      ],
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toEqual([
      { position: 0, color: 'rgba(17,34,51,1)' },
      { position: 1, color: 'rgba(221,238,255,1)' },
    ]);
    expect(rec.rects.some(rect => rect.fs === '[object Object]')).toBe(true);
    expect(rec.rects.some(rect => rect.fs === '#CC0000')).toBe(true);
  });

  it('draws three-series up/down bars and only the explicitly authored marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      stockUpDownBars: true,
      showLegend: true,
      legendPos: 'r',
      series: [
        series({
          name: 'High', values: [55, 57, 57], markerSymbol: 'circle',
          markerSize: 5, markerFill: '4472C4', markerLine: '4472C4', lineHidden: true,
        }),
        series({ name: 'Low', values: [11, 12, 13], markerSymbol: 'none', lineHidden: true }),
        series({ name: 'Close', values: [32, 35, 34], markerSymbol: 'none', lineHidden: true }),
      ],
    }), RECT, 1);

    expect(rec.rects.filter(rect => rect.fs === '#3F3F3F')).toHaveLength(3);
    expect(rec.arcs).toHaveLength(4); // three plotted High points + one legend marker
    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining(['High', 'Low', 'Close']));
    const lastBar = rec.paintEvents.reduce((lastIndex, event, index) =>
      event.kind === 'rect' && event.fillStyle === '#3F3F3F' ? index : lastIndex,
      -1,
    );
    const markerFills = rec.paintEvents.flatMap((event, index) =>
      event.kind === 'fill' && event.fillStyle === '#4472C4' ? [index] : []
    );
    expect(lastBar).toBeGreaterThanOrEqual(0);
    expect(markerFills).toHaveLength(4);
    expect(markerFills.every(index => index > lastBar)).toBe(true);
  });

  it('paints up/down bars over the high-low rule and owning line series', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, stockModel({
      stockUpDownBars: true,
      stockHiLowLineColor: '123456',
      series: [
        series({
          name: 'Open', values: [40], markerSymbol: 'none',
          lineColor: '4472C4', seriesType: 'line',
        }),
        series({
          name: 'Close', values: [20], markerSymbol: 'none',
          lineColor: 'C0504D', seriesType: 'line',
        }),
      ],
      plotGroups: [plotGroup('line', 0, 2)],
    }), RECT, 1);

    const hiLowIndex = rec.paintEvents.findIndex(event =>
      event.kind === 'stroke' && event.strokeStyle === '#123456'
    );
    const owningLineIndex = rec.paintEvents.findIndex(event =>
      event.kind === 'stroke' && event.strokeStyle === '#4472C4'
    );
    const barIndex = rec.paintEvents.findIndex(event =>
      event.kind === 'rect' && event.fillStyle === '#3F3F3F'
    );
    expect(hiLowIndex).toBeGreaterThanOrEqual(0);
    expect(owningLineIndex).toBeGreaterThanOrEqual(0);
    expect(barIndex).toBeGreaterThan(hiLowIndex);
    expect(barIndex).toBeGreaterThan(owningLineIndex);
  });
});

describe('surface contour charts', () => {
  const wireframeSurfaceModel = (over: Partial<ChartModel> = {}): ChartModel => baseModel({
    chartType: 'surface3D',
    categories: ['X1', 'X2'],
    valMin: 0,
    valMax: 10,
    valAxisMajorUnit: 10,
    surfaceWireframe: true,
    catAxisHidden: true,
    valAxisHidden: true,
    catAxisMajorGridlines: false,
    valAxisMajorGridlines: false,
    threeD: {
      rotationX: 30,
      rotationY: 20,
      perspective: 30,
      seriesAxis: { hidden: true, lineHidden: true, majorTickMark: 'none' },
    },
    series: [
      series({ name: 'Y1', values: [2, 4] }),
      series({ name: 'Y2', values: [6, 8] }),
    ],
    ...over,
  });

  it('renders surface3D as the same bounded source-grid mesh without flattening it to another family', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface3D',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 4,
      valAxisMajorUnit: 1,
      surfaceWireframe: false,
      threeD: { rotationX: 30, rotationY: 20, perspective: 30 },
      series: [
        series({ name: 'Y1', values: [1, 2] }),
        series({ name: 'Y2', values: [3, 4] }),
      ],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('Unsupported chart');
    expect(rec.filledPaths.length).toBeGreaterThan(0);
  });

  it('uses the linked dataPointWireframe line for a Surface wireframe mesh', () => {
    const rec = recordingCtx();
    const wireframeGradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: Array.from({ length: 4_096 }, (_, index) => ({
        position: index / 4_095,
        color: index === 0 ? '123456' : 'ABCDEF',
      })),
    };
    const filledSurfaceGradient = {
      ...wireframeGradient,
      stops: [
        { position: 0, color: 'FF0000' },
        { position: 1, color: 'FFCCCC' },
      ],
    };
    renderChart(rec.ctx, wireframeSurfaceModel({
      chartStyleRoles: {
        dataPoint3D: {
          linePaints: [filledSurfaceGradient],
          linePaintAuthored: true,
        },
        dataPointWireframe: {
          linePaints: [wireframeGradient],
          linePaintAuthored: true,
          lineColorIndex: 0,
          lineWidthEmu: 25_400,
          lineDash: 'dash',
          lineCap: 'rnd',
          lineJoin: 'bevel',
        },
      },
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.gradients[0].stops).toHaveLength(4_096);
    expect(rec.gradients[0].stops[0]?.color).toBe('rgba(18,52,86,1)');
    const wireframe = rec.strokeDetails.filter(stroke =>
      stroke.lineWidth === 2
      && stroke.dash.length > 0
      && stroke.cap === 'round'
      && stroke.join === 'bevel'
    );
    expect(wireframe).toHaveLength(4);
  });

  it('adds interpolated band-boundary contours across wireframe cells', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      categories: ['X1', 'X2', 'X3'],
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({ name: 'Y1', values: [0, 0, 0], lineColor: 'AA0000' }),
        series({ name: 'Y2', values: [0, 20, 0] }),
        series({ name: 'Y3', values: [0, 0, 0] }),
      ],
    }), RECT, 1);

    // The uniform line style keeps the 3x3 source mesh at 12 segments. Its
    // four cells each contribute two interpolated contours at value 10.
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '#AA0000'))
      .toHaveLength(20);
  });

  it('keeps unresolved, no-fill, and compound dataPointWireframe lines fail-closed', () => {
    for (const role of [
      { linePaintAuthored: true, lineColorIndex: 0 },
      { lineHidden: true, lineColorIndex: 0 },
      { lineCompound: 'dbl' },
    ]) {
      const rec = recordingCtx();
      renderChart(rec.ctx, wireframeSurfaceModel({
        chartStyleRoles: { dataPointWireframe: role },
      }), RECT, 1);
      expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#595959')).toBe(false);
    }
  });

  it('does not guess a relative dataPointWireframe palette index', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      chartStyleRoles: {
        dataPointWireframe: {
          lineColors: ['123456', 'ABCDEF'],
          linePaintAuthored: true,
        },
      },
    }), RECT, 1);

    expect(rec.strokeDetails.some(stroke =>
      stroke.strokeStyle === '#123456' || stroke.strokeStyle === '#ABCDEF'
    )).toBe(false);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#595959')).toBe(false);
  });

  it('uses the first Surface series as the mesh default and lets direct bands override it', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({
          name: 'Y1', values: [2, 18], lineColor: 'AA0000', lineWidthEmu: 50_800,
          chartexStyle: { lineDash: 'dash', lineDashAuthored: true },
        }),
        series({
          name: 'Y2', values: [2, 18], lineColor: '00AA00', lineWidthEmu: 25_400,
          chartexStyle: { lineDash: 'dot', lineDashAuthored: true },
        }),
      ],
      surfaceBandFormats: [{
        idx: 0,
        lineColor: '00FFFF',
        lineWidthEmu: 38_100,
        style: { lineDash: 'dot', lineDashAuthored: true },
      }],
      chartStyleRoles: {
        dataPointWireframe: {
          lineColors: ['123456'],
          lineColorIndex: 0,
          linePaintAuthored: true,
        },
      },
    }), RECT, 1);

    const mesh = rec.strokeDetails.filter(stroke =>
      stroke.strokeStyle === '#AA0000' || stroke.strokeStyle === '#00FFFF'
    );
    expect(mesh.filter(stroke => stroke.strokeStyle === '#AA0000')).toHaveLength(3);
    expect(mesh.filter(stroke => stroke.strokeStyle === '#00FFFF')).toHaveLength(5);
    expect(mesh.every(stroke => stroke.dash.length > 0)).toBe(true);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#00AA00')).toBe(false);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#123456')).toBe(false);
  });

  it('keeps a uniform first-series wireframe on the source grid and contours', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({
          name: 'Y1', values: [2, 18], lineColor: 'AA0000',
          chartexStyle: { lineDash: 'dash', lineDashAuthored: true },
        }),
        series({ name: 'Y2', values: [2, 18], lineColor: '00AA00' }),
      ],
    }), RECT, 1);

    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '#AA0000'))
      .toHaveLength(6);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle === '#00AA00')).toBe(false);
  });

  it('resolves one structured first-series wireframe recipe for all fallback bands', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({
          name: 'Y1', values: [2, 18],
          chartexStyle: {
            linePaints: [{
              fillType: 'gradient', gradType: 'linear', angle: 0,
              stops: [
                { position: 0, color: '112233' },
                { position: 1, color: 'DDEEFF' },
              ],
            }],
            linePaintAuthored: true,
          },
        }),
        series({ name: 'Y2', values: [2, 18] }),
      ],
      surfaceBandFormats: [{ idx: 0, lineColor: '00FFFF' }],
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(1);
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '[object Object]'))
      .toHaveLength(3);
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '#00FFFF'))
      .toHaveLength(5);
  });

  it('keeps first-series no-line and direct band no-line authoritative per band', () => {
    const firstSeriesNoLine = recordingCtx();
    renderChart(firstSeriesNoLine.ctx, wireframeSurfaceModel({
      series: [
        series({
          name: 'Y1', values: [2, 4],
          chartexStyle: { lineHidden: true, linePaintAuthored: true },
        }),
        series({ name: 'Y2', values: [6, 8], lineColor: '00AA00' }),
      ],
    }), RECT, 1);
    expect(firstSeriesNoLine.strokeDetails.some(stroke =>
      stroke.strokeStyle === '#595959' || stroke.strokeStyle === '#00AA00'
    )).toBe(false);

    const bandNoLine = recordingCtx();
    renderChart(bandNoLine.ctx, wireframeSurfaceModel({
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({ name: 'Y1', values: [2, 18], lineColor: 'AA0000' }),
        series({ name: 'Y2', values: [2, 18], lineColor: '00AA00' }),
      ],
      surfaceBandFormats: [{ idx: 0, lineHidden: true }],
    }), RECT, 1);
    expect(bandNoLine.strokeDetails.filter(stroke => stroke.strokeStyle === '#AA0000'))
      .toHaveLength(3);
    expect(bandNoLine.strokeDetails.some(stroke => stroke.strokeStyle === '#00AA00'))
      .toBe(false);
  });

  it('ignores later-series wireframe defaults and retains automatic band colours', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      valMax: 20,
      valAxisMajorUnit: 10,
      series: [
        series({ name: 'Y1', values: [2, 18] }),
        series({ name: 'Y2', values: [2, 18], lineColor: '00B050' }),
      ],
    }), RECT, 1);

    const meshColors = new Set(rec.strokeDetails
      .map(stroke => stroke.strokeStyle)
      .filter(color => color !== '#000000'));
    expect(meshColors.size).toBe(2);
    expect(meshColors.has('#00B050')).toBe(false);
  });

  it('does not apply dataPointWireframe paint to a filled Surface', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      surfaceWireframe: false,
      chartStyleRoles: {
        dataPointWireframe: {
          linePaints: [{
            fillType: 'gradient',
            gradType: 'linear',
            angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096,
              color: '123456',
            })),
          }],
          linePaintAuthored: true,
          lineColorIndex: 0,
        },
      },
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(0);
    expect(rec.filledPaths.length).toBeGreaterThan(0);
  });

  it('rejects an oversized dataPointWireframe line before resolving it', () => {
    const oversized = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: Array.from({ length: 4_097 }, (_, index) => ({
        position: index / 4_096,
        color: '123456',
      })),
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      chartStyleRoles: {
        dataPointWireframe: {
          linePaints: [oversized],
          linePaintAuthored: true,
          lineColorIndex: 0,
        },
      },
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(0);
    expect(rec.strokeDetails).toHaveLength(0);

    const direct = recordingCtx();
    renderChart(direct.ctx, wireframeSurfaceModel({
      surfaceBandFormats: [{
        idx: 0,
        style: { linePaints: [oversized], linePaintAuthored: true },
      }],
    }), RECT, 1);
    expect(direct.gradients).toHaveLength(0);
    expect(direct.strokeDetails).toHaveLength(0);
  });

  it('does not charge or resolve a first-series recipe fully overridden by direct bands', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, wireframeSurfaceModel({
      series: [
        series({
          name: 'Y1', values: [2, 4],
          chartexStyle: {
            linePaints: [{
              fillType: 'gradient',
              gradType: 'linear',
              angle: 0,
              stops: Array.from({ length: 4_097 }, (_, index) => ({
                position: index / 4_096,
                color: '123456',
              })),
            }],
            linePaintAuthored: true,
          },
        }),
        series({ name: 'Y2', values: [6, 8] }),
      ],
      surfaceBandFormats: [{ idx: 0, lineColor: '00FFFF' }],
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(0);
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle === '#00FFFF'))
      .toHaveLength(4);
  });

  it('resolves Surface band paint once with direct structured/unresolved/no-fill precedence', () => {
    const rec = recordingCtx();
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: [
        { position: 0, color: '112233' },
        { position: 1, color: 'DDEEFF' },
      ],
    };
    const directGradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 90,
      stops: [
        { position: 0, color: 'AA0000' },
        { position: 1, color: 'FFCCCC' },
      ],
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 40,
      valAxisMajorUnit: 10,
      surfaceWireframe: false,
      chartStyleRoles: {
        dataPoint3D: { fillPaints: [gradient], fillPaintAuthored: true, lineHidden: true },
      },
      surfaceBandFormats: [
        {
          idx: 0,
          style: { fillHidden: true, fillPaintAuthored: true },
          fillHidden: true,
        },
        { idx: 1, style: { fillPaintAuthored: true } },
        {
          idx: 2,
          style: { fillPaints: [directGradient], fillPaintAuthored: true },
        },
      ],
      series: [
        series({ name: 'Y1', values: [5, 35] }),
        series({ name: 'Y2', values: [35, 5] }),
      ],
    }), RECT, 1);

    // Band 0 is direct noFill, band 1 is authored-but-unresolved, band 2 uses
    // its direct recipe, and band 3 uses the linked role. Each visible recipe
    // is registered once despite spanning several clipped polygons.
    expect(rec.gradients).toHaveLength(2);
    expect(rec.gradients.map(item => item.stops[0]?.color)).toEqual([
      'rgba(170,0,0,1)',
      'rgba(17,34,51,1)',
    ]);
    expect(rec.filledPaths.length).toBeGreaterThan(1);
  });

  it('paints Surface3D wall roles and keeps direct unresolved/no-fill authoritative', () => {
    const rec = recordingCtx();
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 90,
      stops: [
        { position: 0, color: '334455' },
        { position: 1, color: 'CCDDEE' },
      ],
    };
    renderChart(rec.ctx, baseModel({
      chartType: 'surface3D',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      surfaceWireframe: false,
      chartStyleRoles: {
        floor: { fillPaints: [gradient], fillPaintAuthored: true },
        wall: { fillPaints: [gradient], fillPaintAuthored: true },
      },
      threeD: {
        rotationX: 30,
        rotationY: 20,
        perspective: 30,
        floor: { fillHidden: true },
        backWall: { style: { fillPaintAuthored: true } },
      },
      series: [
        series({ name: 'Y1', values: [2, 4] }),
        series({ name: 'Y2', values: [6, 8] }),
      ],
    }), RECT, 1);

    // floor=noFill and authored-but-unresolved backWall both suppress linked
    // paint. Only sideWall consumes the linked wall gradient.
    expect(rec.gradients).toHaveLength(1);
  });

  it('uses the same authored wall-thickness slabs for Surface3D', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface3D',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      surfaceWireframe: false,
      threeD: {
        rotationX: 20,
        rotationY: 20,
        perspective: 30,
        floor: { fillColor: 'FF0000', thicknessPercent: 25 },
        sideWall: { fillColor: '00B050', thicknessPercent: 25 },
        backWall: { fillColor: 'AA00FF', thicknessPercent: 25 },
      },
      series: [
        series({ name: 'Y1', values: [2, 4] }),
        series({ name: 'Y2', values: [6, 8] }),
      ],
    }), RECT, 1);
    for (const color of ['#FF0000', '#00B050', '#AA00FF']) {
      expect(rec.filledPaths.filter(path => path.fillStyle === color).length).toBeGreaterThan(1);
    }
  });

  it('continues Surface3D category and value gridlines across visible thick slab faces', () => {
    const render = (thicknessPercent: number) => {
      const rec = strokedPolylineCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'surface3D',
        categories: ['X1', 'X2', 'X3'],
        catAxisCrossBetween: 'between',
        catAxisMajorGridlines: true,
        catAxisGridlineColor: 'FF00FF',
        catAxisGridlineWidthEmu: 25_400,
        catAxisGridlineDash: 'dash',
        catAxisMinorGridlines: true,
        catAxisMinorGridlineColor: 'FF8000',
        catAxisMinorGridlineWidthEmu: 25_400,
        catAxisMinorGridlineDash: 'dot',
        valAxisMajorGridlines: true,
        valAxisGridlineColor: '00FFFF',
        valAxisGridlineWidthEmu: 25_400,
        valAxisGridlineDash: 'dot',
        valAxisMinorGridlines: true,
        valAxisMinorGridlineColor: '123456',
        valAxisMinorGridlineWidthEmu: 25_400,
        valAxisMinorGridlineDash: 'dash',
        valAxisMinorUnit: 1,
        valMin: 0,
        valMax: 10,
        valAxisMajorUnit: 2,
        surfaceWireframe: false,
        threeD: {
          rotationX: 20, rotationY: 20, perspective: 30,
          floor: { fillColor: 'C00000', thicknessPercent },
          sideWall: { fillColor: '008000', thicknessPercent },
          backWall: { fillColor: '4472C4', thicknessPercent },
        },
        series: [
          series({ name: 'Y1', values: [2, 5, 3] }),
          series({ name: 'Y2', values: [4, 9, 6] }),
        ],
      }), RECT, 1);
      return rec.strokes;
    };
    const planar = render(0);
    const thick = render(25);
    for (const color of ['#FF00FF', '#FF8000', '#00FFFF', '#123456']) {
      const planarLines = planar.filter(stroke => stroke.ss === color);
      const thickLines = thick.filter(stroke => stroke.ss === color);
      expect(planarLines.length).toBeGreaterThan(0);
      expect(thickLines.length).toBeGreaterThan(planarLines.length);
      expect(thickLines.every(stroke => stroke.lw === 2)).toBe(true);
      expect(thickLines.every(stroke => stroke.dash.length > 0)).toBe(true);
    }
  });

  it('rejects an oversized linked Surface recipe before resolving any paint', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      surfaceWireframe: false,
      chartStyleRoles: {
        dataPoint3D: {
          fillPaints: [{
            fillType: 'gradient',
            gradType: 'linear',
            angle: 0,
            stops: Array.from({ length: 4_097 }, (_, index) => ({
              position: index / 4_096,
              color: '112233',
            })),
          }],
          fillPaintAuthored: true,
        },
      },
      series: [
        series({ name: 'Y1', values: [2, 4] }),
        series({ name: 'Y2', values: [6, 8] }),
      ],
    }), RECT, 1);

    expect(rec.gradients).toHaveLength(0);
    expect(rec.filledPaths).toHaveLength(0);
  });

  it('centres category points but places Surface series on axis endpoints', () => {
    const rec = strokedPolylineCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      catAxisCrossBetween: 'between',
      catAxisLineColor: '00FFFF',
      catAxisMajorGridlines: true,
      catAxisGridlineColor: '00AA00',
      valMin: 0,
      valMax: 30,
      valAxisMajorUnit: 5,
      valAxisHidden: true,
      surfaceWireframe: false,
      threeD: {
        rotationX: 15,
        rotationY: 20,
        rightAngleAxes: true,
        perspective: 0,
        seriesAxis: {
          hidden: false,
          lineHidden: false,
          lineColor: 'FF00FF',
          majorTickMark: 'none',
        },
      },
      series: [
        series({ name: 'Y1', values: [0, 30] }),
        series({ name: 'Y2', values: [30, 0] }),
      ],
    }), RECT, 1);

    const fractionsAlong = (axisColor: string, labels: string[], labelOffset: { x: number; y: number }) => {
      const axis = rec.strokes.find(stroke => stroke.ss === axisColor && stroke.points.length === 2);
      expect(axis).toBeDefined();
      const [start, end] = axis!.points;
      const dx = end.x - start.x;
      const dy = end.y - start.y;
      const lengthSquared = dx * dx + dy * dy;
      return labels.map(label => {
      const text = rec.texts.find(entry => entry.text === label);
      expect(text).toBeDefined();
        const point = { x: text!.x - labelOffset.x, y: text!.y - labelOffset.y };
      return ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared;
      });
    };
    expect(fractionsAlong('#FF00FF', ['Y1', 'Y2'], { x: 8, y: 0 }))
      .toEqual([expect.closeTo(0, 6), expect.closeTo(1, 6)]);
    expect(fractionsAlong('#00FFFF', ['X1', 'X2'], { x: 0, y: 8 }))
      .toEqual([expect.closeTo(0.25, 6), expect.closeTo(0.75, 6)]);

    const categoryAxis = rec.strokes.find(stroke =>
      stroke.ss === '#00FFFF' && stroke.points.length === 2
    );
    expect(categoryAxis).toBeDefined();
    const [categoryStart, categoryEnd] = categoryAxis!.points;
    const categoryDx = categoryEnd.x - categoryStart.x;
    const categoryDy = categoryEnd.y - categoryStart.y;
    const categoryLengthSquared = categoryDx * categoryDx + categoryDy * categoryDy;
    const categoryGridStrokes = rec.strokes
      .filter(stroke => stroke.ss === '#00AA00' && stroke.points.length === 2);
    // Excel continues each planar rule across both the floor and back wall.
    // Positive-thickness exterior continuation has its own face-count tests.
    expect(categoryGridStrokes).toHaveLength(6);
    const gridFractions = categoryGridStrokes
      .filter((_, index) => index % 2 === 0)
      .map(stroke => {
        const point = stroke.points[0];
        return ((point.x - categoryStart.x) * categoryDx
          + (point.y - categoryStart.y) * categoryDy) / categoryLengthSquared;
      });
    expect(gridFractions).toEqual([
      expect.closeTo(0, 6), expect.closeTo(0.5, 6), expect.closeTo(1, 6),
    ]);
  });

  it('uses the observed oblique perspective camera when view3D is omitted', () => {
    const model = (threeD: ChartModel['threeD']): ChartModel => baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2', 'X3'],
      valMin: 0,
      valMax: 30,
      valAxisMajorUnit: 5,
      surfaceWireframe: false,
      threeD,
      series: [
        series({ name: 'Y1', values: [0, 10, 20] }),
        series({ name: 'Y2', values: [10, 20, 30] }),
      ],
    });
    const paintPoints = (threeD: ChartModel['threeD']) => {
      const rec = recordingCtx();
      renderChart(rec.ctx, model(threeD), RECT, 1);
      return rec.filledPaths.map(path => path.points);
    };

    const omitted = paintPoints(undefined);
    const observedDefault = paintPoints({
      rotationX: 15, rotationY: 20, rightAngleAxes: false, perspective: 30,
    });
    const orthographic = paintPoints({
      rotationX: 15, rotationY: 20, rightAngleAxes: true, perspective: 0,
    });
    expect(omitted).toEqual(observedDefault);
    expect(omitted).not.toEqual(orthographic);
    // The S1-S5 Office vectors show stronger vertical convergence than the
    // normative pinhole response alone. These normalized mesh extrema pin the
    // bounded Surface-family perspective gain without coupling other 3-D
    // families to the compatibility observation.
    const points = omitted.flat();
    expect(Math.min(...points.map(point => point.x))).toBeCloseTo(149.2, 0);
    expect(Math.max(...points.map(point => point.y))).toBeCloseTo(264.4, 0);
  });

  it('rejects an unbounded authored band count before allocating tick arrays', () => {
    const rec = recordingCtx();
    expect(() => renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 1,
      valAxisMajorUnit: Number.MIN_VALUE,
      surfaceWireframe: false,
      series: [
        series({ name: 'Y1', values: [0, 1] }),
        series({ name: 'Y2', values: [1, 0] }),
      ],
    }), RECT, 1)).not.toThrow();
    expect(rec.filledPaths).toEqual([]);

    const finiteButExcessive = recordingCtx();
    expect(() => renderChart(finiteButExcessive.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 1,
      valAxisMajorUnit: 1e-5,
      surfaceWireframe: false,
      series: [
        series({ name: 'Y1', values: [0, 1] }),
        series({ name: 'Y2', values: [1, 0] }),
      ],
    }), RECT, 1)).not.toThrow();
    expect(finiteButExcessive.filledPaths).toEqual([]);
  });

  it('does not extend the observed automatic material to an unverified camera', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      surfaceWireframe: false,
      legacyChartStyle: 2,
      themeAccentColors: ['156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'],
      threeD: { rotationX: 30, rotationY: 45, perspective: 20, rightAngleAxes: false },
      series: [
        series({ name: 'Y1', values: [5, 5] }),
        series({ name: 'Y2', values: [5, 5] }),
      ],
    }), RECT, 1);
    expect(rec.filledPaths.filter(path => path.points.length >= 3).map(path => path.fillStyle))
      .toEqual(expect.arrayContaining(['#156082']));
    expect(rec.filledPaths.filter(path => path.points.length >= 3)
      .every(path => path.fillStyle === '#156082')).toBe(true);
  });

  it('selects the upper Surface diagonal across plane and reversed saddle boundaries', () => {
    const diagonalDirection = (values: number[][]): number => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'surface',
        categories: ['X1', 'X2'],
        valMin: 0,
        valMax: 30,
        valAxisMajorUnit: 30,
        surfaceWireframe: false,
        threeD: {
          rotationX: 90,
          rotationY: 0,
          perspective: 0,
          rightAngleAxes: false,
          seriesAxis: { hidden: true, majorTickMark: 'none', lineHidden: true },
        },
        series: values.map((row, index) => series({ name: `Y${index + 1}`, values: row })),
      }), RECT, 1);

      const triangles = rec.filledPaths.filter(path => path.points.length === 3);
      expect(triangles).toHaveLength(2);
      const shared = triangles[0].points.filter(point => triangles[1].points.some(other =>
        Math.abs(point.x - other.x) < 1e-9 && Math.abs(point.y - other.y) < 1e-9
      ));
      expect(shared).toHaveLength(2);
      return (shared[1].x - shared[0].x) * (shared[1].y - shared[0].y);
    };

    // Equal opposing sums use the stable B-D tie direction. Reversing a
    // saddle swaps which opposing pair is higher and therefore swaps the
    // selected diagonal instead of turning Excel's ridge into a valley.
    expect(diagonalDirection([[0, 10], [20, 30]])).toBeGreaterThan(0);
    expect(diagonalDirection([[0, 30], [30, 0]])).toBeGreaterThan(0);
    expect(diagonalDirection([[30, 0], [0, 30]])).toBeLessThan(0);
  });

  it('uses the automatic Pattern 2 band palette and view-dependent legend order', () => {
    const model = (rotationX: number): ChartModel => baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 0,
      valMax: 30,
      valAxisMajorUnit: 5,
      showLegend: true,
      legendPos: 'b',
      surfaceWireframe: false,
      legacyChartStyle: 2,
      themeAccentColors: ['156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'],
      threeD: {
        rotationX,
        seriesAxis: { hidden: true, majorTickMark: 'none', lineHidden: true },
      },
      series: [
        series({ name: 'Y1', values: [0, 10] }),
        series({ name: 'Y2', values: [20, 30] }),
      ],
    });
    const oblique = recordingCtx();
    renderChart(oblique.ctx, model(15), RECT, 1);
    const contour = recordingCtx();
    renderChart(contour.ctx, model(90), RECT, 1);

    const palette = ['#115473', '#CF642B', '#155E1F', '#0C8CBD', '#8E2582', '#449428'];
    const legendColors = new Set(oblique.rects.map(rect => rect.fs));
    expect(palette.every(color => legendColors.has(color))).toBe(true);
    const obliqueBands = oblique.texts.map(text => text.text).filter(text => text.includes('-'));
    const contourBands = contour.texts.map(text => text.text).filter(text => text.includes('-'));
    expect(obliqueBands).toEqual(['0-5', '5-10', '10-15', '15-20', '20-25', '25-30']);
    expect(contourBands).toEqual(['25-30', '20-25', '15-20', '10-15', '5-10', '0-5']);
  });

  it('derives automatic Surface bands from the projected value-axis length', () => {
    const surface = (
      values: number[][],
      threeD: ChartModel['threeD'],
      rect: ChartRect,
    ): string[] => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'surface',
        categories: values[0].map((_, index) => `X${index + 1}`),
        showLegend: true,
        legendPos: 'b',
        surfaceWireframe: false,
        threeD: {
          ...threeD,
          seriesAxis: { hidden: true, majorTickMark: 'none', lineHidden: true },
        },
        series: values.map((row, index) => series({ name: `Y${index + 1}`, values: row })),
      }), rect, 1);
      return rec.texts.map(text => text.text).filter(text => text.includes('-'));
    };

    expect(surface([
      [10, 20, 30, 20, 10],
      [20, 40, 60, 40, 20],
      [30, 60, 90, 60, 30],
      [20, 40, 60, 40, 20],
      [10, 20, 30, 20, 10],
    ], { rotationX: 90, rotationY: 0, perspective: 0, rightAngleAxes: false }, RECT))
      .toEqual(['80-100', '60-80', '40-60', '20-40', '0-20']);

    expect(surface([
      [0, 5, 10, 15],
      [10, 15, 20, 25],
      [20, 25, 30, 35],
    ], {}, { x: 0, y: 0, w: 900, h: 220 }))
      .toEqual(['0-10', '10-20', '20-30', '30-40']);
  });

  it('anchors automatic Surface band boundaries at an authored non-zero minimum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2'],
      valMin: 3,
      valAxisMajorUnit: 5,
      showLegend: true,
      legendPos: 'b',
      surfaceWireframe: false,
      series: [
        series({ name: 'Y1', values: [4, 8] }),
        series({ name: 'Y2', values: [9, 12] }),
      ],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text).filter(text => text.includes('-')))
      .toEqual(['3-8', '8-13']);
  });

  it('interpolates value bands across the matrix and labels both category axes', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2', 'X3'],
      valMin: 0,
      valMax: 100,
      valAxisMajorUnit: 20,
      showLegend: true,
      legendPos: 'r',
      surfaceWireframe: false,
      threeD: { seriesAxis: { hidden: false, orientation: 'minMax', majorTickMark: 'cross', lineHidden: false } },
      series: [
        series({ name: 'Y1', color: '156082', values: [10, 20, 10] }),
        series({ name: 'Y2', color: 'E97132', values: [20, 40, 20] }),
        series({ name: 'Y3', color: '196B24', values: [30, 60, 30] }),
      ],
    }), RECT, 1);

    expect(rec.filledPaths.length).toBeGreaterThan(8);
    const surfaceFills = [...new Set(rec.filledPaths.map(path => path.fillStyle))];
    for (const base of ['156082', 'E97132', '196B24']) {
      expect(surfaceFills.some(fill => isSurfaceMaterialColor(fill, base))).toBe(true);
    }
    expect(surfaceFills.length).toBeGreaterThan(3);
    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining(['X1', 'X2', 'X3', 'Y1', 'Y2', 'Y3', '0-20', '80-100']));
  });

  it('paints authored major and minor ticks on both ordinal Surface axes', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2', 'X3', 'X4', 'X5'],
      valMin: 0,
      valMax: 100,
      valAxisHidden: true,
      catAxisLineColor: 'AA0000',
      catAxisLineHidden: false,
      catAxisMajorTickMark: 'cross',
      catAxisMinorTickMark: 'cross',
      threeD: {
        rotationX: 90,
        rotationY: 0,
        perspective: 0,
        rightAngleAxes: false,
        seriesAxis: {
          hidden: false,
          orientation: 'minMax',
          majorTickMark: 'cross',
          minorTickMark: 'cross',
          lineColor: '0000AA',
          lineHidden: false,
        },
      },
      series: [
        series({ name: 'Y1', values: [10, 20, 30, 20, 10] }),
        series({ name: 'Y2', values: [20, 40, 60, 40, 20] }),
        series({ name: 'Y3', values: [30, 60, 90, 60, 30] }),
        series({ name: 'Y4', values: [20, 40, 60, 40, 20] }),
        series({ name: 'Y5', values: [10, 20, 30, 20, 10] }),
      ],
    }), RECT, 1);

    const categoryTicks = rec.segs.filter(segment =>
      segment.ss === '#AA0000'
      && Math.abs(segment.x1 - segment.x0) < 0.01
      && Math.abs(segment.y1 - segment.y0) <= 6.01
    );
    expect(categoryTicks.filter(segment =>
      Math.abs(Math.abs(segment.y1 - segment.y0) - 6) < 0.01
    )).toHaveLength(5);
    expect(categoryTicks.filter(segment =>
      Math.abs(Math.abs(segment.y1 - segment.y0) - 4) < 0.01
    )).toHaveLength(4);

    const seriesTicks = rec.segs.filter(segment =>
      segment.ss === '#0000AA'
      && Math.abs(segment.y1 - segment.y0) < 0.01
      && Math.abs(segment.x1 - segment.x0) <= 6.01
    );
    expect(seriesTicks.filter(segment =>
      Math.abs(Math.abs(segment.x1 - segment.x0) - 6) < 0.01
    )).toHaveLength(5);
    expect(seriesTicks.filter(segment =>
      Math.abs(Math.abs(segment.x1 - segment.x0) - 4) < 0.01
    )).toHaveLength(4);
  });

  it('uses the shared automatic value-axis unit with one bounded surface material', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'surface',
      categories: ['X1', 'X2', 'X3', 'X4', 'X5'],
      showLegend: true,
      legendPos: 'r',
      surfaceWireframe: false,
      valAxisHidden: true,
      threeD: {
        rotationX: 90,
        rotationY: 0,
        perspective: 0,
        rightAngleAxes: false,
        seriesAxis: {
          hidden: false, orientation: 'minMax', majorTickMark: 'cross', lineHidden: false,
        },
      },
      series: [
        series({ name: 'Y1', color: '156082', values: [10, 20, 30, 20, 10] }),
        series({ name: 'Y2', color: 'E97132', values: [20, 40, 60, 40, 20] }),
        series({ name: 'Y3', color: '196B24', values: [30, 60, 90, 60, 30] }),
        series({ name: 'Y4', color: '0F9ED5', values: [20, 40, 60, 40, 20] }),
        series({ name: 'Y5', color: 'A02B93', values: [10, 20, 30, 20, 10] }),
      ],
    }), RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels.some(label => /^0-/.test(label))).toBe(true);
    expect(labels.some(label => /-100$/.test(label))).toBe(true);
    const automaticColors = new Set([
      '#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93',
      '#4472C4', '#ED7D31', '#A9D18E', '#FF0000', '#70AD47', '#4BACC6',
      '#FFC000', '#9E480E', '#843C0C', '#636363', '#255E91', '#967300',
    ]);
    expect(rec.filledPaths.every(path => [...automaticColors].some(base =>
      isSurfaceMaterialColor(path.fillStyle, base)
    ))).toBe(true);
  });
});

// ─── CH14 — pie callout data labels (Word boxed labels, §21.2.2.197) ─────────
//
// When a pie/doughnut series `<c:dLbls>` carries a `<c:spPr>` box shape the
// labels are drawn as boxed callouts OUTSIDE each slice: a filled+bordered
// rectangle with the category name and percent on separate lines, plus a
// leader line back to the rim for a box pulled far from its slice. Without a
// box shape the historical plain-text label path is preserved.

/** A pie model whose series data labels request Word's boxed callout layout. */
function pieCalloutModel(over: Partial<ChartModel> = {}): ChartModel {
  return baseModel({
    chartType: 'pie',
    categories: ['Brazil', 'Vietnam', 'Colombia', 'Indonesia', 'Honduras', 'Other'],
    series: [series({
      name: 'Prod',
      values: [51500, 28500, 14000, 10800, 8349, 61000],
      seriesDataLabels: {
        showVal: false, showCatName: true, showSerName: false, showPercent: true,
        position: 'bestFit',
        labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
        showLeaderLines: true,
        leaderLineColor: 'A6A6A6',
        leaderLineWidthEmu: 9525,
      },
      dataLabelOverrides: [
        // idx 0 (Brazil) is a per-point styling override: empty text (reuses the
        // composed cat/percent), blue font, its own box.
        { idx: 0, text: '', position: 'bestFit', fontColor: '4472C4', fontSizeHpt: 1000, fontBold: false,
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 } },
      ],
    })],
    ...over,
  });
}

describe('CH14 — pie callout data labels', () => {
  it('keeps border-only bestFit labels on their slices without inventing callout leaders', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Point 1', 'Point 2', 'Point 3', 'Point 4'],
      series: [series({
        values: [8, 6, 2, 2],
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          showLeaderLines: true,
          leaderLineColor: '808080',
          labelBox: {
            fillHidden: true,
            fillPaintAuthored: true,
            borderColor: '808080',
            borderWidthEmu: 3175,
          },
        },
      })],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(expect.arrayContaining(['8', '6', '2']));
    expect(rec.strokeRects.some(rect => rect.ss === '#808080')).toBe(true);
    expect(rec.strokedPaths.filter(path => path.strokeStyle === '#808080')).toHaveLength(0);
  });

  it('paints a custom rich callout from the same measured inline block', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Only'],
      themeMajorFontLatin: 'Major Theme',
      themeMinorFontLatin: 'Minor Theme',
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0,
          text: 'Rich Callout',
          richRuns: [
            { text: 'Rich', fontFace: '+mj-lt', color: '112233' },
            { text: ' Callout', fontFace: '+mn-lt', color: '445566' },
          ],
          manualLayout: { xMode: 'edge', yMode: 'edge', x: 0.25, y: 0.25, w: 0.2, h: 0.1 },
          labelBox: { fill: 'ABCDEF', borderColor: '123456' },
        }],
      })],
    }), RECT, 1);

    expect(rec.rects.some(box => box.fs === '#ABCDEF' && box.w === 128 && box.h === 36))
      .toBe(true);
    expect(rec.texts.find(call => call.text === 'Rich'))
      .toMatchObject({ fillStyle: '#112233' });
    expect(rec.texts.find(call => call.text === 'Rich')?.font)
      .toContain('"Major Theme"');
    expect(rec.texts.find(call => call.text === ' Callout'))
      .toMatchObject({ fillStyle: '#445566' });
  });

  it('enters rich pie label layout for a per-point label without series defaults', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Only'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0,
          text: 'point only',
          showVal: false,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          labelBox: { fill: 'ABCDEF' },
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toContain('point only');
    expect(rec.rects.some(rect => rect.fs === '#ABCDEF')).toBe(true);
  });

  it.each(['ctr', 'inEnd'])(
    'keeps a point-only callout at its authored %s anchor without boxing siblings',
    position => {
      const rec = recordingCtx();
      renderChart(rec.ctx, baseModel({
        chartType: 'pie',
        categories: ['A', 'B'],
        series: [series({
          values: [1, 1],
          seriesDataLabels: {
            showVal: false,
            showCatName: true,
            showSerName: false,
            showPercent: false,
            position: 'outEnd',
          },
          dataLabelOverrides: [{
            idx: 0,
            text: 'A',
            position,
            labelBox: { fill: 'ABCDEF' },
          }],
        })],
      }), RECT, 1);

      const boxes = rec.rects.filter(rect => rect.fs === '#ABCDEF');
      expect(boxes).toHaveLength(1);
      const box = boxes[0];
      const center = rec.arcs[0];
      const outerR = Math.max(...rec.arcs.map(arc => arc.r));
      const boxDistance = Math.hypot(
        box.x + box.w / 2 - center.x,
        box.y + box.h / 2 - center.y,
      );
      expect(boxDistance).toBeLessThan(outerR);
      const sibling = rec.texts.find(text => text.text === 'B');
      expect(sibling).toBeDefined();
      expect(Math.hypot(sibling!.x - center.x, sibling!.y - center.y)).toBeGreaterThan(outerR);
    },
  );

  it('composes a per-point show flag without series label defaults', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Only'],
      series: [series({ values: [1], dataLabelOverrides: [{ idx: 0, text: '', showPercent: true }] })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).toContain('100%');
  });

  it('applies a per-point manual pie layout without series label defaults', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Only'],
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0,
          text: 'manual point',
          manualLayout: { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5, w: 0.2, h: 0.1 },
        }],
      })],
    }), RECT, 1);
    expect(rec.texts.find(text => text.text === 'manual point')).toMatchObject({ x: 384, y: 198 });
  });

  it('keeps a per-point delete from falling back to legacy pie percentages', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      showDataLabels: true,
      categories: ['Only'],
      series: [series({ values: [1], dataLabelOverrides: [{ idx: 0, text: '', deleted: true }] })],
    }), RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('100%');
  });

  it('anchors an outer-ring doughnut label in that ring rather than the full doughnut band', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'doughnut',
      holeSize: 50,
      categories: ['A', 'B'],
      series: [
        series({
          name: 'Outer',
          values: [1, 1],
          dataLabelOverrides: [{ idx: 0, text: 'outer-ring' }],
        }),
        series({ name: 'Inner', values: [1, 1] }),
      ],
    }), RECT, 1);

    const radii = [...new Set(rec.arcs.map(arc => arc.r))].sort((a, b) => b - a);
    const outerLabel = rec.fontTexts.find(text => text.text === 'outer-ring');
    expect(radii.length).toBeGreaterThanOrEqual(3);
    expect(outerLabel).toBeDefined();
    const centerX = rec.arcs[0].x;
    expect(outerLabel!.x - centerX).toBeCloseTo((radii[0] + radii[1]) / 2, 3);
  });

  it('draws a filled callout box per slice (category name + percent on separate lines)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieCalloutModel(), RECT, 1);
    // White box fills — one per drawn label. The box fill is the parsed white
    // (`#FFFFFF`); the slice wedges fill with palette colors, so filter on the
    // box color to isolate the callout rectangles.
    const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
    expect(boxes.length).toBe(6);
    // Category names and percents are drawn as SEPARATE fillText lines.
    const texts = rec.texts.map(t => t.text);
    expect(texts).toContain('Brazil');
    expect(texts).toContain('Other');
    expect(texts).toContain('30%'); // 51500 / 174149 ≈ 29.6% → 30
    expect(texts).toContain('16%'); // 28500 / 174149 ≈ 16.4% → 16
    // No space-joined "Brazil 30%" composite — category and percent are split.
    expect(texts.some(t => /Brazil\s+\d/.test(t))).toBe(false);
  });

  it('colors the per-point (Brazil) label with its override font color', () => {
    // Purpose-built context that snapshots fillStyle with each fillText so the
    // per-point font-color override (`#4472C4` for Brazil vs `#000` default)
    // can be asserted directly.
    const calls: { text: string; fs: string }[] = [];
    const state: Record<string, unknown> = {
      font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    };
    const handler: ProxyHandler<Record<string, unknown>> = {
      get(_t, prop: string) {
        if (prop in state && typeof state[prop] !== 'function') return state[prop];
        if (prop === 'measureText') return (t: string) => ({ width: String(t).length * 6 });
        if (prop === 'fillText') return (text: string) => calls.push({ text, fs: String(state.fillStyle) });
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
        return () => undefined;
      },
      set(_t, prop: string, value) { state[prop] = value; return true; },
    };
    const ctx = new Proxy(state, handler) as unknown as CanvasRenderingContext2D;
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    const brazil = calls.find(c => c.text === 'Brazil');
    expect(brazil?.fs).toBe('#4472C4');
    // A non-overridden slice uses the default black font (no series fontColor).
    const other = calls.find(c => c.text === 'Other');
    expect(other?.fs).toBe('#000');
  });

  it('draws leader lines in the parsed leader color when a box is pulled off its slice', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, pieCalloutModel(), RECT, 1);
    // Leader lines are stroked in the parsed leader color (#A6A6A6). The small
    // slices (Colombia/Indonesia/Honduras) get pulled far enough out to draw a
    // leader; assert at least one leader segment exists in that color.
    const leaders = rec.segs.filter(s => s.ss === '#A6A6A6');
    expect(leaders.length).toBeGreaterThan(0);
  });

  it('applies authored manual width and height to callout paint and text bounds', () => {
    const rec = recordingCtx();
    const model = pieCalloutModel();
    model.series[0].dataLabelOverrides = [{
      idx: 0,
      text: 'A long manually sized callout label',
      manualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.25, y: 0.25, w: 0.2, h: 0.1,
      },
      labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
    }];
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.rects.some(box =>
      box.fs === '#FFFFFF'
      && Math.abs(box.x - 160) < 1e-6
      && Math.abs(box.y - 90) < 1e-6
      && Math.abs(box.w - 128) < 1e-6
      && Math.abs(box.h - 36) < 1e-6
    )).toBe(true);
    expect(rec.texts.some(text => text.text.includes('…'))).toBe(true);
  });

  it('keeps plain-text labels (no boxes) when the dLbls carries no box shape', () => {
    const rec = recordingCtx();
    const model = pieCalloutModel();
    // Strip the box → falls back to the historical plain outer-ring text path.
    const sdl = model.series[0].seriesDataLabels;
    if (sdl) { sdl.labelBox = undefined; sdl.showLeaderLines = false; }
    model.series[0].dataLabelOverrides = null;
    renderChart(rec.ctx, model, RECT, 1);
    // No white callout boxes are drawn.
    expect(rec.rects.filter(r => r.fs === '#FFFFFF').length).toBe(0);
  });

  // #767 — the bestFit de-overlap must keep every callout box INSIDE the chart
  // rect even when many slivers stack in one column. The old separate() slid the
  // column up for a bottom overflow, then unconditionally slid it back down for a
  // top underflow, cancelling the up-slide so a 9+-label column spilled ~200px
  // past the bottom edge. Stress with many same-side slivers and assert 0
  // overflow + 0 overlap.
  function pieStressModel(): ChartModel {
    // 14 slices: 12 slivers swept early (clustered top→right→bottom, so most
    // land in ONE column) + 2 large slices. Single-line percent labels keep each
    // box short enough that a 9+-box column still fits within the plot band, so
    // both invariants (0 overflow AND 0 overlap) can hold — the regime the old
    // cancel-slide broke by spilling the column ~200px past the bottom edge.
    const cats: string[] = [];
    const values: number[] = [];
    for (let i = 0; i < 12; i++) { cats.push(`Sliver ${i + 1}`); values.push(3); }
    cats.push('Big A'); values.push(40);
    cats.push('Big B'); values.push(40);
    return baseModel({
      chartType: 'pie',
      title: 'Coffee Production',
      categories: cats,
      series: [series({
        name: 'Prod',
        values,
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: true,
          position: 'bestFit',
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
        },
      })],
    });
  }

  it('keeps every callout box inside the chart rect with no overlaps under many slivers (#767)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieStressModel(), RECT, 1);
    const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
    // All 14 slices are drawn (none dropped).
    expect(boxes.length).toBe(14);

    // (a) 0 overflow: every box lies fully within the chart rect [0..h] × [0..w].
    for (const b of boxes) {
      expect(b.y).toBeGreaterThanOrEqual(RECT.y - 0.5);
      expect(b.y + b.h).toBeLessThanOrEqual(RECT.y + RECT.h + 0.5);
      expect(b.x).toBeGreaterThanOrEqual(RECT.x - 0.5);
      expect(b.x + b.w).toBeLessThanOrEqual(RECT.x + RECT.w + 0.5);
    }

    // The stress must actually pack 9+ boxes into a single vertical column — the
    // regime that broke the old cancel-slide. Split by column via box centre x.
    const midX = RECT.x + RECT.w / 2;
    const rightCol = boxes.filter(b => b.x + b.w / 2 >= midX)
      .sort((p, q) => p.y - q.y);
    const leftCol = boxes.filter(b => b.x + b.w / 2 < midX)
      .sort((p, q) => p.y - q.y);
    expect(Math.max(rightCol.length, leftCol.length)).toBeGreaterThanOrEqual(9);

    // (b) 0 overlap within each column: consecutive boxes never overlap
    // vertically (boxes in different columns may share a y-band harmlessly).
    for (const col of [rightCol, leftCol]) {
      for (let k = 1; k < col.length; k++) {
        expect(col[k].y).toBeGreaterThanOrEqual(col[k - 1].y + col[k - 1].h - 0.5);
      }
    }
  });

  // #767 (follow-up) — the original stress above used SINGLE-line percent labels,
  // whose short boxes never triggered the TOP-underflow half of the bug. The old
  // separate() slid a bottom-heavy column UP to clear the bottom edge, then failed
  // to slide it back DOWN: its cap measured "room" against the bottom edge the
  // up-slide had just pinned (room = 0), so a top underflow of ~40-100px was left
  // uncorrected. The guard was ASYMMETRIC — it kept boxes off the BOTTOM but let
  // the first box of an up-slid column escape well ABOVE the plot top.
  //
  // This case reproduces the top escape with TALL two-line labels (long wrapped
  // category name + percent, showCatName + showPercent) and bottom-heavy slice
  // orders that pack many slivers into ONE column at the pie's BOTTOM — before the
  // symmetric round-trip clamp these drove the topmost box to y ≈ -40…-100. It
  // asserts 0 overflow at BOTH the top AND the bottom edge across several
  // geometries and slice arrangements.
  //
  // Overlap is deliberately NOT asserted here: with this many two-line boxes the
  // column genuinely over-packs (more label than the plot can hold), so the
  // documented over-pack path lets boxes touch/overlap rather than escape the
  // frame — trading escape for overlap is the whole point of the clamp. The
  // 0-overlap invariant is covered by the single-line stress above, whose short
  // boxes DO fit, which is exactly why that case uses single-line labels.
  function pieTwoLineStressModel(
    arrange: 'bottomHeavy' | 'bigMid',
    firstSliceAngle: number,
  ): ChartModel {
    const cats: string[] = [];
    const values: number[] = [];
    const longName = (i: number): string => `Very Long Category Name Number ${i + 1}`;
    if (arrange === 'bottomHeavy') {
      // A big slice, then 12 slivers, then a big slice — the slivers sweep
      // through the bottom into one column, each label two lines tall.
      cats.push('Big A'); values.push(48);
      for (let i = 0; i < 12; i++) { cats.push(longName(i)); values.push(3); }
      cats.push('Big B'); values.push(48);
    } else {
      // Big slices in the middle of the order rotate the sliver run to the top
      // half, another arrangement that drove the pre-fix top escape.
      for (let i = 0; i < 5; i++) { cats.push(longName(i)); values.push(3); }
      cats.push('Big A'); values.push(40);
      cats.push('Big B'); values.push(40);
      for (let i = 5; i < 10; i++) { cats.push(longName(i)); values.push(3); }
    }
    return baseModel({
      chartType: 'pie',
      title: 'Coffee Production',
      categories: cats,
      firstSliceAngle,
      series: [series({
        name: 'Prod',
        values,
        seriesDataLabels: {
          // TWO lines per label: category name + percent → a tall box, the regime
          // the single-line stress above never reached.
          showVal: false, showCatName: true, showSerName: false, showPercent: true,
          position: 'bestFit',
          labelBox: { fill: 'FFFFFF', borderColor: '4472C4', borderWidthEmu: 12700 },
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
        },
      })],
    });
  }

  // Geometries + slice arrangements that all drove a top-edge escape before the
  // symmetric round-trip clamp. Each combo must keep every box inside the plot
  // rect at BOTH ends.
  const stressGeoms: Array<[string, ChartRect]> = [
    ['tall', { x: 0, y: 0, w: 640, h: 360 }],
    ['square', { x: 0, y: 0, w: 400, h: 400 }],
    ['wide', { x: 0, y: 0, w: 700, h: 300 }],
  ];
  const stressCases: Array<['bottomHeavy' | 'bigMid', number]> = [
    ['bottomHeavy', 0],
    ['bigMid', 180],
  ];
  for (const [gName, geom] of stressGeoms) {
    for (const [arrange, fsa] of stressCases) {
      it(`two-line callouts stay inside the rect at BOTH edges (${gName}/${arrange}) (#767)`, () => {
        const rec = recordingCtx();
        renderChart(rec.ctx, pieTwoLineStressModel(arrange, fsa), geom, 1);
        const boxes = rec.rects.filter(r => r.fs === '#FFFFFF');
        // Every slice's callout is drawn (none dropped).
        expect(boxes.length).toBeGreaterThanOrEqual(12);

        // (a) 0 overflow at the TOP edge — the half of #767 the old guard missed
        // (pre-fix this drove the topmost box to a negative y, ~40-100px above
        // the plot top).
        for (const b of boxes) {
          expect(b.y, `top overflow in ${gName}/${arrange}`).toBeGreaterThanOrEqual(geom.y - 0.5);
        }
        // (a') 0 overflow at the BOTTOM edge — the half #767 already guarded.
        for (const b of boxes) {
          expect(b.y + b.h, `bottom overflow in ${gName}/${arrange}`).toBeLessThanOrEqual(geom.y + geom.h + 0.5);
        }
        // Horizontal containment stays intact too.
        for (const b of boxes) {
          expect(b.x).toBeGreaterThanOrEqual(geom.x - 0.5);
          expect(b.x + b.w).toBeLessThanOrEqual(geom.x + geom.w + 0.5);
        }

        // The stress must actually pack a deep single column — the regime that
        // broke the old cancel-slide. Split by box centre x.
        const midX = geom.x + geom.w / 2;
        const rightCol = boxes.filter(b => b.x + b.w / 2 >= midX);
        const leftCol = boxes.filter(b => b.x + b.w / 2 < midX);
        expect(Math.max(rightCol.length, leftCol.length)).toBeGreaterThanOrEqual(6);
      });
    }
  }
});

// CH15 — chartEx box-and-whisker (MS 2014 chartex ext). Verify the derived
// statistics (exclusive quartiles + 1.5·IQR outlier fence + mean) and the
// value-axis scale drive observable geometry: the IQR box rects, the outlier
// dots, and the nice-rounded axis labels.
describe('CH15 — chartEx box-and-whisker', () => {
  // The sample-24 Category-1 orange series: an obvious outlier at 128 sits far
  // beyond Q3 + 1.5·IQR, so the whisker stops at 34 and 128 is drawn as a dot.
  const CAT1_ORANGE = [-3, 1, -6, 10, 34, 128, 22, -12, -28];

  function boxModel(over: Partial<ChartModel> = {}): ChartModel {
    return baseModel({
      chartType: 'boxWhisker',
      title: 'box',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexBox: {
        categories: ['Category 1'],
        series: [
          {
            name: 'S1',
            color: 'ED7D31',
            valuesByCategory: [CAT1_ORANGE],
            meanMarker: true,
            meanLine: false,
            showOutliers: true,
            showNonoutliers: false,
            quartileMethod: 'exclusive',
          },
        ],
      },
      ...over,
    });
  }

  it('keeps ChartEx value-label offset stable across a larger authored font', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 10,
      valAxisFontSizeHpt: 1200,
      valAxisLineColor: '123456',
      valAxisLineHidden: false,
    }), RECT, 1);

    const axis = rec.segs.find(segment =>
      segment.ss === '#123456'
      && Math.abs(segment.x0 - segment.x1) < 0.001
      && Math.abs(segment.y1 - segment.y0) > 100
    );
    const zero = rec.texts.find(text => text.text === '0' && text.align === 'right');
    expect(axis).toBeDefined();
    expect(zero).toBeDefined();
    expect(axis!.x0 - zero!.x).toBeCloseTo(7, 5);
  });

  it('preserves the authored box-series outline on the filled legend key', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      legendFontSizeHpt: 1500,
      chartexBox: {
        categories: ['6'],
        series: [{
          name: 'Super Duper MPG',
          color: 'FF00FF',
          lineColor: '000000',
          lineWidthEmu: 12700,
          valuesByCategory: [[10, 20, 30]],
          meanMarker: true,
          meanLine: false,
          showOutliers: true,
          showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);

    const key = rec.rects.find(rect =>
      rect.fs === '#FF00FF' && Math.abs(rect.w - 7) < 0.01 && Math.abs(rect.h - 7) < 0.01
    );
    expect(key).toBeDefined();
    expect(rec.strokeRects.some(rect =>
      rect.ss === '#000000'
      && rect.lw === 1
      && Math.abs(rect.w - 6) < 0.01
      && Math.abs(rect.h - 6) < 0.01
    )).toBe(true);
  });

  it('wraps a long side-legend series name into the measured two-line band', () => {
    // Approximate a narrow Office theme face: the first two words fit the
    // bounded side column while the complete quoted name does not.
    const rec = recordingCtx((text, fontPx) => Array.from(text).length * fontPx * 0.48);
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      legendFontSizeHpt: 1500,
      chartexBox: {
        categories: ['6'],
        series: [{
          name: '"Super Duper MPG"',
          color: 'FF00FF',
          valuesByCategory: [[10, 20, 30]],
          meanMarker: true,
          meanLine: false,
          showOutliers: true,
          showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), { x: 0, y: 0, w: 494, h: 288 }, 4 / 3);

    const legendText = rec.texts
      .filter(text => text.text.includes('Super') || text.text.includes('MPG'))
      .map(text => text.text);
    expect(legendText).toEqual(['"Super Duper', 'MPG"']);
    expect(legendText.some(text => text.includes('…'))).toBe(false);
  });

  it('keeps Chart Style NoStyle distinct from an explicit noFill on the box legend key', () => {
    const noStyle = recordingCtx();
    renderChart(noStyle.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: { lineHidden: true, lineNoStyle: true },
    }), RECT, 1);
    expect(noStyle.strokeRects.some(rect =>
      Math.abs(rect.w - 6) < 0.01 && Math.abs(rect.h - 6) < 0.01
    )).toBe(true);

    const explicitNoFill = recordingCtx();
    renderChart(explicitNoFill.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: { lineHidden: true },
    }), RECT, 1);
    expect(explicitNoFill.strokeRects.some(rect =>
      Math.abs(rect.w - 6) < 0.01 && Math.abs(rect.h - 6) < 0.01
    )).toBe(false);
  });

  it('labels the value axis with Excel nice-rounded gridline values including a negative bound', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel(), RECT, 1);
    const labels = rec.texts.map(t => t.text);
    // The data spans −28..128, so the auto axis must reach BELOW zero (a
    // negative label) and ABOVE the max (a label ≥ 128's rounded ceiling),
    // and cross zero. Exact bounds depend on the axis length, so assert the
    // scale SHAPE rather than pinned numbers.
    expect(labels).toContain('0');
    expect(labels.some(l => l.startsWith('-'))).toBe(true);
    expect(labels.some(l => Number(l) >= 130)).toBe(true);
  });

  it.each([
    { name: 'wide range', values: [10, 466], rect: { x: 0, y: 0, w: 371, h: 216 }, step: 50, max: 500 },
    { name: 'ordinary range', values: [0, 27], rect: { x: 0, y: 0, w: 530, h: 396 }, step: 5, max: 30 },
    { name: 'fence boundary', values: [0, 12.0001], rect: { x: 0, y: 0, w: 530, h: 396 }, step: 2, max: 14 },
  ])('uses the compact Office-observed automatic box axis for $name', ({ values, rect, step, max }) => {
    const rec = markerRecordingCtx();
    const model = boxModel({
      title: null,
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S1',
          color: 'ED7D31',
          valuesByCategory: [values],
          meanMarker: false,
          meanLine: false,
          showOutliers: false,
          showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    });
    renderChart(rec.ctx, model, rect, 1);

    const ticks = rec.texts
      .map(text => Number(text.text))
      .filter(Number.isFinite)
      .sort((left, right) => left - right);
    expect(ticks[0]).toBe(0);
    expect(ticks.at(-1)).toBe(max);
    expect(ticks[1] - ticks[0]).toBeCloseTo(step, 8);
  });

  it('draws the authored value-axis title, rule, gridline style, and explicit 0.2 major unit', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 1,
      valMax: 3,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisTitle: 'A&R readiness score',
      valAxisTitleFontSizeHpt: 900,
      valAxisTitleFontBold: false,
      valAxisTitleFontColor: '404040',
      valAxisLineColor: 'BFBFBF',
      valAxisLineWidthEmu: 9525,
      valAxisMajorTickMark: 'out',
      valAxisMajorGridlines: true,
      valAxisGridlineColor: 'D9D9D9',
      valAxisGridlineWidthEmu: 9525,
    }), RECT, 1);

    expect(rec.texts.some(text => text.text === 'A&R readiness score')).toBe(true);
    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['1.0', '1.2', '1.4', '1.6', '1.8', '2.0', '2.2', '2.4', '2.6', '2.8', '3.0']),
    );
    expect(rec.segs.some(segment =>
      Math.abs(segment.x0 - segment.x1) < 0.5 &&
      Math.abs(segment.y1 - segment.y0) > 100 &&
      segment.ss.toLowerCase() === '#bfbfbf'
    )).toBe(true);
    expect(rec.segs.filter(segment =>
      Math.abs(segment.y0 - segment.y1) < 0.5 &&
      Math.abs(segment.x1 - segment.x0) > 100 &&
      segment.ss.toLowerCase() === '#d9d9d9'
    ).length).toBeGreaterThanOrEqual(10);
  });

  it('derives the omitted box-axis major unit from authored min/max bounds', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 1,
      valMax: 3,
      valAxisMajorUnit: null,
      valAxisFormatCode: '0.0',
    }), RECT, 1);

    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['1.0', '1.2', '1.4', '1.6', '1.8', '2.0', '2.2', '2.4', '2.6', '2.8', '3.0']),
    );
  });

  it('uses the authored ChartEx value-axis font size for numeric tick labels', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 1,
      valMax: 3,
      valAxisMajorUnit: 0.2,
      valAxisFormatCode: '0.0',
      valAxisFontSizeHpt: 900,
      valAxisFontFace: 'Calibri',
      valAxisFontItalic: true,
      valAxisTitle: 'Miles per Gallon',
      valAxisTitleFontSizeHpt: 1200,
      valAxisTitleFontBold: false,
      valAxisTitleFontItalic: true,
    }), RECT, 1);

    const tick = rec.fontTexts.find(text => text.text === '1.0');
    expect(tick).toBeDefined();
    expect(tick?.font).toContain('9px');
    expect(tick?.font).toContain('Calibri');
    expect(tick?.font).toContain('italic');
    const title = rec.fontTexts.find(text => text.text === 'Miles per Gallon');
    expect(title?.font).toContain('italic');
    expect(title?.font).not.toContain('bold');
  });

  it('uses the authored category-axis rule and keeps both axis labels clear of cross ticks', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 0,
      valMax: 500,
      valAxisMajorUnit: 50,
      valAxisFormatCode: '0.0',
      valAxisFontSizeHpt: 1200,
      valAxisLineColor: '000000',
      valAxisLineWidthEmu: 12700,
      valAxisMajorTickMark: 'cross',
      catAxisFontSizeHpt: 1000,
      catAxisLineColor: '000000',
      catAxisLineWidthEmu: 12700,
      catAxisMajorTickMark: 'cross',
    }), RECT, 1);

    const horizontalAxis = rec.segs.find(segment =>
      Math.abs(segment.y0 - segment.y1) < 0.5 &&
      Math.abs(segment.x1 - segment.x0) > 100 &&
      segment.ss.toLowerCase() === '#000000'
    );
    const verticalAxis = rec.segs.find(segment =>
      Math.abs(segment.x0 - segment.x1) < 0.5 &&
      Math.abs(segment.y1 - segment.y0) > 100 &&
      segment.ss.toLowerCase() === '#000000'
    );
    expect(horizontalAxis).toBeDefined();
    expect(verticalAxis).toBeDefined();
    expect(horizontalAxis?.lw).toBe(1);

    const categoryTicks = rec.segs.filter(segment =>
      Math.abs(segment.x0 - segment.x1) < 0.5 &&
      // Office's authored cross tick length is 6pt total (3pt per side), not
      // 6pt on both sides of the axis.
      Math.abs(segment.y1 - segment.y0) >= 5.5 &&
      Math.abs((segment.y0 + segment.y1) / 2 - (horizontalAxis?.y0 ?? 0)) < 0.5 &&
      segment.ss.toLowerCase() === '#000000'
    );
    expect(categoryTicks.length).toBeGreaterThanOrEqual(1);

    const categoryLabel = rec.texts.find(text => text.text === 'Category 1');
    expect(categoryLabel).toBeDefined();
    expect((categoryLabel?.y ?? 0) - (horizontalAxis?.y0 ?? 0)).toBeGreaterThanOrEqual(8);

    const zeroLabel = rec.texts.find(text => text.text === '0.0');
    expect(zeroLabel).toBeDefined();
    expect((verticalAxis?.x0 ?? 0) - (zeroLabel?.x ?? 0)).toBeCloseTo(7);
  });

  it('places the value axis from measured tick-label width instead of a fixed chart-width gutter', () => {
    const axisX = (formatCode: string): number => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, boxModel({
        valMin: 1,
        valMax: 3,
        valAxisMajorUnit: 0.2,
        valAxisFormatCode: formatCode,
        valAxisFontSizeHpt: 900,
        valAxisTitle: 'Score',
        valAxisTitleFontSizeHpt: 900,
        valAxisLineColor: 'BFBFBF',
        valAxisLineWidthEmu: 9525,
      }), RECT, 1);
      const axis = rec.segs.find(segment =>
        Math.abs(segment.x0 - segment.x1) < 0.5 &&
        Math.abs(segment.y1 - segment.y0) > 100 &&
        segment.ss.toLowerCase() === '#bfbfbf'
      );
      if (!axis) throw new Error('value axis not drawn');
      return axis.x0;
    };

    const shortLabelsAxisX = axisX('0.0');
    const longLabelsAxisX = axisX('0.00000000');
    expect(longLabelsAxisX).toBeGreaterThan(shortLabelsAxisX + 20);
  });

  it('draws exactly one IQR box rect and one outlier dot for a single box with one outlier', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel(), RECT, 1);
    // Exactly one filled IQR rect (Q1..Q3) for the single box.
    expect(rec.fillRects.length).toBe(1);
    // The 128 point is the sole outlier → one dot (arc). The box-and-whisker
    // renderer draws arcs ONLY for outliers (the mean `×` and whiskers are line
    // segments), so the arc count equals the outlier count.
    expect(rec.arcs.length).toBe(1);
    // The outlier dot sits ABOVE the box top (smaller y = higher value).
    const box = rec.fillRects[0];
    expect(rec.arcs[0].y).toBeLessThan(box.y);
  });

  it('draws every non-outlier sample point when the visibility flag is enabled', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: null, valuesByCategory: [CAT1_ORANGE],
          meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    // Eight interior points plus the single outlier at 128.
    expect(rec.arcs.length).toBe(CAT1_ORANGE.length);
  });

  it('fills box-and-whisker sample points when the marker role uses Chart Style NoStyle', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointMarkerStyle: {
        fillHidden: true,
        fillNoStyle: true,
        lineHidden: true,
        lineNoStyle: true,
      },
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(CAT1_ORANGE.length);
    expect(rec.fillCalls).toBe(CAT1_ORANGE.length);
  });

  it('keeps box-and-whisker sample points transparent for an explicit marker noFill', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointMarkerStyle: { fillHidden: true, lineHidden: true },
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: null, valuesByCategory: [CAT1_ORANGE],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(CAT1_ORANGE.length);
    expect(rec.fillCalls).toBe(0);
  });

  it('does not draw box sample points when the Chart Style marker symbol is none', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexMarkerSymbol: 'none',
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.arcs).toHaveLength(0);
  });

  it('keeps a series-local noFill line suppressed over Chart Style NoStyle', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointStyle: { lineHidden: true, lineNoStyle: true },
      chartexDataPointLineStyle: { lineHidden: true, lineNoStyle: true },
      chartexDataPointMarkerStyle: { lineHidden: true, lineNoStyle: true },
      valAxisHidden: true,
      catAxisHidden: true,
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'exclusive', chartexStyle: { lineHidden: true },
        }],
      },
    }), RECT, 1);
    expect(rec.segs).toHaveLength(0);
  });

  it('draws authored box-and-whisker minor ticks', () => {
    const count = (minorTick: ChartModel['valAxisMinorTickMark']): number => {
      const rec = segRecordingCtx();
      renderChart(rec.ctx, boxModel({
        valMin: 3, valMax: 23, valAxisMajorUnit: 10, valAxisMinorUnit: 4,
        valAxisMajorGridlines: false, valAxisMinorGridlines: false,
        valAxisMajorTickMark: 'none', valAxisMinorTickMark: minorTick,
      }), RECT, 1);
      return rec.segs.filter(segment =>
        Math.abs(segment.y1 - segment.y0) < 0.01
        && Math.abs(segment.x1 - segment.x0) > 0
        && Math.abs(segment.x1 - segment.x0) <= 12
      ).length;
    };
    expect(count('cross') - count('none')).toBe(4);
  });

  it('draws 6pt major ticks and shorter 4pt minor ticks', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 3,
      valMax: 23,
      valAxisMajorUnit: 10,
      valAxisMinorUnit: 4,
      valAxisMajorGridlines: false,
      valAxisMinorGridlines: false,
      valAxisMajorTickMark: 'out',
      valAxisMinorTickMark: 'out',
      valAxisLineColor: 'FF00FF',
      valAxisLineWidthEmu: 9525,
    }), RECT, 2);

    const tickLengths = rec.segs
      .filter(segment =>
        segment.ss.toLowerCase() === '#ff00ff'
        && Math.abs(segment.y1 - segment.y0) < 0.01
        && Math.abs(segment.x1 - segment.x0) <= 20
      )
      .map(segment => Math.abs(segment.x1 - segment.x0));
    expect(tickLengths).toContain(12);
    expect(tickLengths).toContain(8);
  });

  it('uses Excel box-and-whisker semantic sizes for sample dots and the mean marker', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexMarkerSizePt: 6,
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [[1, 2, 3]],
          meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 2);

    expect(rec.arcs).toHaveLength(3);
    // Office's vector output uses 3pt-diameter observation dots even though
    // the linked Chart Style's generic marker layout says size=6 here.
    expect(rec.arcs.every(arc => arc.r === 3)).toBe(true);
    const meanCross = rec.segments.find(segment =>
      segment.length === 4
      && segment.every((point, index) => index === 0
        || Math.abs(point.x - segment[0].x) > 0
        || Math.abs(point.y - segment[0].y) > 0)
    );
    expect(meanCross).toBeDefined();
    // The semantic mean marker is a fixed 6pt square (12px at 2px/pt).
    expect(Math.abs((meanCross as Array<{ x: number; y: number }>)[1].x - (meanCross as Array<{ x: number; y: number }>)[0].x)).toBe(12);
    expect(Math.abs((meanCross as Array<{ x: number; y: number }>)[1].y - (meanCross as Array<{ x: number; y: number }>)[0].y)).toBe(12);
    expect(Math.abs((meanCross as Array<{ x: number; y: number }>)[3].x - (meanCross as Array<{ x: number; y: number }>)[2].x)).toBe(12);
    expect(Math.abs((meanCross as Array<{ x: number; y: number }>)[3].y - (meanCross as Array<{ x: number; y: number }>)[2].y)).toBe(12);
  });

  it('uses the authored Chart Style marker symbol for box sample points', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexMarkerSizePt: 6,
      chartexMarkerSymbol: 'square',
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(0);
    // One IQR box plus three square raw-point markers.
    expect(rec.fillRects.length).toBeGreaterThanOrEqual(4);
  });

  it('uses median-of-halves quartiles for inclusive and exclusive methods', () => {
    const values = [1, 2, 3, 4, 100];
    const exclusive = markerRecordingCtx();
    renderChart(exclusive.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [values],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    const inclusive = markerRecordingCtx();
    renderChart(inclusive.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [values],
          meanMarker: false, meanLine: false, showOutliers: true, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);

    // Inclusive includes the median in each half: Q3=4, so 100 is an outlier.
    // Exclusive omits it: Q3=(4+100)/2, so the same point stays inside.
    expect(exclusive.arcs).toHaveLength(0);
    expect(inclusive.arcs).toHaveLength(1);
  });

  it('suppresses outlier dots when <cx:visibility outliers="0">', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['Category 1'],
        series: [{
          name: 'S1', color: 'ED7D31', valuesByCategory: [CAT1_ORANGE],
          meanMarker: true, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.arcs.length).toBe(0);
  });

  it('draws one IQR box per (category, series) — 3 categories × 2 series = 6 boxes', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A', 'B', 'C'],
        series: [
          { name: 'S1', color: '5B9BD5', valuesByCategory: [[1, 2, 3], [4, 5, 6], [7, 8, 9]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: false, quartileMethod: 'exclusive' },
          { name: 'S2', color: 'ED7D31', valuesByCategory: [[2, 3, 4], [5, 6, 7], [8, 9, 10]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: false, quartileMethod: 'exclusive' },
        ],
      },
    }), RECT, 1);
    expect(rec.fillRects.length).toBe(6);
  });

  it('lays out formula-only one-box-per-series data as full slots with a legend', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      catAxisHidden: true,
      chartexBox: {
        oneBoxPerSeries: true,
        categories: ['Foundations', 'Adaptation'],
        series: [
          { name: 'Foundations', color: '5B9BD5', valuesByCategory: [[1, 2, 3], []], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true, quartileMethod: 'exclusive' },
          { name: 'Adaptation', color: 'ED7D31', valuesByCategory: [[], [4, 5, 6]], meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true, quartileMethod: 'inclusive' },
        ],
      },
    }), RECT, 1);

    const boxes = rec.fillRects.filter(rect => rect.w > 80);
    expect(boxes).toHaveLength(2);
    const orderedBoxes = [...boxes].sort((a, b) => a.x - b.x);
    const centerDistance = (orderedBoxes[1].x + orderedBoxes[1].w / 2)
      - (orderedBoxes[0].x + orderedBoxes[0].w / 2);
    // Formula-only ChartEx stores each visible box as a separate series in one
    // category group. catScaling@gapWidth applies around that whole group, not
    // independently between the diagonal series entries.
    expect(orderedBoxes[0].w / centerDistance).toBeCloseTo(
      1 - BOX_WHISKER_SLOT_GUTTER_FRACTION,
      5,
    );
    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['Foundations', 'Adaptation']),
    );
  });

  it('places the first and last box half a category interval from the plot edges', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      title: null,
      valMin: 0,
      valMax: 10,
      valAxisMajorUnit: 2,
      valAxisMajorGridlines: true,
      chartexBox: {
        categories: ['A', 'B', 'C'],
        series: [{
          name: 'S', color: '5B9BD5', valuesByCategory: [[1, 2, 3], [4, 5, 6], [7, 8, 9]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);

    const horizontalSegments = rec.segments.flatMap(segment =>
      segment.slice(1).map((point, index) => ({
        x0: segment[index].x,
        y0: segment[index].y,
        x1: point.x,
        y1: point.y,
      })),
    ).filter(segment => Math.abs(segment.y0 - segment.y1) < 0.5);
    const plotRule = horizontalSegments.sort((a, b) =>
      Math.abs(b.x1 - b.x0) - Math.abs(a.x1 - a.x0)
    )[0];
    if (!plotRule) throw new Error('plot gridline not drawn');

    const centers = rec.fillRects
      .map(rect => rect.x + rect.w / 2)
      .sort((a, b) => a - b);
    expect(centers).toHaveLength(3);
    const interval = centers[1] - centers[0];
    const plotLeft = Math.min(plotRule.x0, plotRule.x1);
    const plotRight = Math.max(plotRule.x0, plotRule.x1);
    expect(centers[0] - plotLeft).toBeCloseTo(interval / 2, 5);
    expect(plotRight - centers[2]).toBeCloseTo(interval / 2, 5);
  });

  it('strokes the box outline with the resolved per-accent ChartEx data-point line', () => {
    const rec = segRecordingCtx();
    const model = boxModel();
    const lineStyle = { lineColors: ['BE6427'], lineWidthEmu: 9525 };
    model.chartexDataPointStyle = lineStyle;
    model.chartexDataPointLineStyle = lineStyle;
    model.chartexDataPointMarkerStyle = lineStyle;
    renderChart(rec.ctx, model, RECT, 1);
    const accentSegs = rec.segs.filter(s => s.ss.toLowerCase() === '#be6427');
    // median + two whisker stems + two whisker caps + mean × (2 strokes) = ≥5
    // accent-colored segments (gridlines/axis use gray, not the accent).
    expect(accentSegs.length).toBeGreaterThanOrEqual(5);
    // The un-darkened fill accent must never be a stroke color.
    expect(rec.segs.some(s => s.ss.toLowerCase() === '#ed7d31')).toBe(false);
  });

  it('uses the specified base-color mapping without inventing linear brightness', () => {
    const rec = recordingCtx();
    const values = [[1, 2, 3]];
    renderChart(rec.ctx, boxModel({
      chartexColorStyleMethod: 'acrossLinear',
      chartexColorPalette: ['FF0000', '00FF00', '0000FF'],
      chartexBox: {
        categories: ['A'],
        series: ['A', 'B', 'C'].map(name => ({
          name,
          color: null,
          valuesByCategory: values,
          meanMarker: false,
          meanLine: false,
          showOutliers: false,
          showNonoutliers: false,
          quartileMethod: 'inclusive',
        })),
      },
    }), RECT, 1);
    const boxFills = rec.rects.map(rect => rect.fs.toUpperCase());
    // acrossLinear selects by relative index. MS-ODRAWXML does not define the
    // brightness range/color space, so the authored colors remain unchanged.
    expect(boxFills).toEqual(['#FF0000', '#00FF00', '#0000FF']);
  });

  it('uses dataPointLine for mean connectors and keeps dataPoint paint separate', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointStyle: { fillColors: ['F4B183'], lineColors: ['C00000'] },
      chartexDataPointLineStyle: { lineColors: ['0070C0'], lineWidthEmu: 25400 },
      chartexBox: {
        categories: ['A', 'B'],
        series: [{
          name: 'S', color: null, valuesByCategory: [[1, 2, 3], [4, 5, 6]],
          meanMarker: false, meanLine: true, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    const lineRole = rec.segs.filter(segment => segment.ss.toLowerCase() === '#0070c0');
    expect(lineRole.some(segment => Math.abs(segment.x1 - segment.x0) > 100)).toBe(true);
    expect(lineRole.every(segment => segment.lw === 2)).toBe(true);
  });

  it('partitions one category into equal series slots with a fixed gutter', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A'],
        series: [
          {
            name: 'S1', color: '5B9BD5', valuesByCategory: [[1, 2, 3, 4]],
            meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
            quartileMethod: 'inclusive',
          },
          {
            name: 'S2', color: 'ED7D31', valuesByCategory: [[2, 3, 4, 5]],
            meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
            quartileMethod: 'inclusive',
          },
        ],
      },
    }), RECT, 1);

    expect(rec.rects).toHaveLength(2);
    expect(rec.rects[0].w).toBeCloseTo(rec.rects[1].w, 5);
    const gutter = rec.rects[1].x - (rec.rects[0].x + rec.rects[0].w);
    const slotWidth = rec.rects[1].x - rec.rects[0].x;
    expect(gutter / slotWidth).toBeCloseTo(0.06, 5);
  });

  it('keeps each series in its stable slot when peer categories are empty', () => {
    const sparse = recordingCtx();
    const model = boxModel({
      chartexBox: {
        categories: ['A', 'B'],
        series: [
          {
            name: 'S1', color: '5B9BD5', valuesByCategory: [[1, 2, 3, 4], []],
            meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
            quartileMethod: 'inclusive',
          },
          {
            name: 'S2', color: 'ED7D31', valuesByCategory: [[], [2, 3, 4, 5]],
            meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
            quartileMethod: 'inclusive',
          },
        ],
      },
    });
    renderChart(sparse.ctx, model, RECT, 1);

    const full = recordingCtx();
    const fullModel = structuredClone(model);
    if (!fullModel.chartexBox) throw new Error('box fixture missing');
    fullModel.chartexBox.series[0].valuesByCategory[1] = [1, 2, 3, 4];
    fullModel.chartexBox.series[1].valuesByCategory[0] = [2, 3, 4, 5];
    renderChart(full.ctx, fullModel, RECT, 1);

    expect(sparse.rects).toHaveLength(2);
    expect(full.rects).toHaveLength(4);
    expect(sparse.rects[0].x).toBeCloseTo(full.rects[0].x, 5);
    expect(sparse.rects[1].x).toBeCloseTo(full.rects[3].x, 5);
    expect(sparse.rects[0].w).toBeCloseTo(sparse.rects[1].w, 5);
  });

  it('drops non-finite observations before deriving or painting geometry', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: '5B9BD5', valuesByCategory: [[NaN, -Infinity, 5, 5, Infinity, 5]],
          meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.rects).toHaveLength(1);
    expect(rec.rects.every(rect => [rect.x, rect.y, rect.w, rect.h].every(Number.isFinite))).toBe(true);
  });

  it('does not emit non-finite Canvas geometry for extreme finite observations', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: '5B9BD5',
          valuesByCategory: [[-Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE]],
          meanMarker: true, meanLine: false, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'exclusive',
        }],
      },
    }), RECT, 1);
    const coordinates = [
      ...rec.fillRects.flatMap(rect => [rect.x, rect.y, rect.w, rect.h]),
      ...rec.arcs.flatMap(arc => [arc.x, arc.y, arc.r]),
      ...rec.segments.flatMap(segment => segment.flatMap(point => [point.x, point.y])),
      ...rec.texts.flatMap(text => [text.x, text.y]),
    ];
    expect(coordinates.length).toBeGreaterThan(0);
    expect(coordinates.every(Number.isFinite)).toBe(true);
  });

  it('atomically rejects finite input whose automatic nice-axis headroom overflows', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: '5B9BD5', valuesByCategory: [[-8e307, 0, 8e307]],
          meanMarker: true, meanLine: true, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.fillRects).toHaveLength(0);
    expect(rec.arcs).toHaveLength(0);
    expect(rec.segments).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toEqual(['(chart values out of range)']);
  });

  it('atomically rejects a huge finite authored range with a tiny major unit', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, boxModel({
      valMin: 0,
      valMax: 1e308,
      valAxisMajorUnit: 1,
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: '5B9BD5', valuesByCategory: [[0, 1, 2]],
          meanMarker: true, meanLine: true, showOutliers: true, showNonoutliers: true,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.fillRects).toHaveLength(0);
    expect(rec.arcs).toHaveLength(0);
    expect(rec.segments).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toEqual(['(chart values out of range)']);
  });

  it('refuses an oversized box-and-whisker chart before sorting or painting', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: '5B9BD5', valuesByCategory: [new Array(10_001).fill(1)],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.rects).toHaveLength(0);
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
  });

  it('uses the box/whisker semantic outline for Chart Style NoStyle only', () => {
    const noStyle = segRecordingCtx();
    renderChart(noStyle.ctx, boxModel({
      chartexDataPointStyle: { lineHidden: true, lineNoStyle: true },
      chartexDataPointLineStyle: { lineHidden: true, lineNoStyle: true },
      chartexDataPointMarkerStyle: { lineHidden: true, lineNoStyle: true },
    }), RECT, 1);
    expect(noStyle.segs.some(segment => segment.ss.toLowerCase() === '#ed7d31')).toBe(true);

    const explicitNoFill = segRecordingCtx();
    renderChart(explicitNoFill.ctx, boxModel({
      chartexDataPointStyle: { lineHidden: true },
      chartexDataPointLineStyle: { lineHidden: true },
      chartexDataPointMarkerStyle: { lineHidden: true },
    }), RECT, 1);
    expect(explicitNoFill.segs.some(segment => segment.ss.toLowerCase() === '#ed7d31')).toBe(false);
  });

  it('lets an explicit series line override a hidden mean-line style', () => {
    const rec = segRecordingCtx();
    renderChart(rec.ctx, boxModel({
      chartexDataPointLineStyle: { lineHidden: true },
      chartexBox: {
        categories: ['A', 'B'],
        series: [{
          name: 'S', color: null, lineColor: 'C00000', lineWidthEmu: 25400,
          valuesByCategory: [[1, 2, 3], [4, 5, 6]],
          meanMarker: false, meanLine: true, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.segs.some(segment =>
      segment.ss.toLowerCase() === '#c00000'
      && Math.abs(segment.x1 - segment.x0) > 100
      && segment.lw === 2
    )).toBe(true);
  });

  it('uses the effective ChartEx fill for both a box and its legend marker', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: { fillColors: ['8064A2'] },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'Styled', color: null, valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);
    expect(rec.rects.filter(rect => rect.fs.toUpperCase() === '#8064A2').length).toBeGreaterThanOrEqual(2);
  });

  it('uses one shared DrawingML gradient recipe for a box and legend and honors rotWithShape', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      showLegend: true,
      legendPos: 'r',
      chartexDataPointStyle: {
        fillPaints: [{
          fillType: 'gradient',
          stops: [
            { position: 0, color: '4472C4' },
            { position: 1, color: 'FFFFFF' },
          ],
          angle: 90,
          gradType: 'linear',
          rotWithShape: false,
        }],
      },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'Gradient', color: null, valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1, 30);
    expect(rec.rects.filter(rect => rect.fs === '[object Object]').length).toBeGreaterThanOrEqual(2);
    expect(rec.gradients.length).toBeGreaterThanOrEqual(2);
    for (const gradient of rec.gradients) {
      const [x1, y1, x2, y2] = gradient.args;
      const dx = x2 - x1;
      const dy = y2 - y1;
      // The host frame is already rotated 30°. rotWithShape=false therefore
      // counter-rotates the authored 90° gradient to 60° in local coordinates.
      expect(dx).toBeGreaterThan(0);
      expect(dy / dx).toBeCloseTo(Math.sqrt(3), 5);
      expect(gradient.stops).toEqual([
        { position: 0, color: 'rgba(68,114,196,1)' },
        { position: 1, color: 'rgba(255,255,255,1)' },
      ]);
    }
  });

  it('honors an explicit ChartEx series outline color and width', () => {
    const rec = segRecordingCtx();
    const model = boxModel();
    const firstSeries = model.chartexBox?.series[0];
    if (!firstSeries) throw new Error('box series fixture missing');
    firstSeries.lineColor = '404040';
    firstSeries.lineWidthEmu = 25400;
    renderChart(rec.ctx, model, RECT, 1);
    const outlineSegments = rec.segs.filter(segment => segment.ss.toLowerCase() === '#404040');
    expect(outlineSegments.length).toBeGreaterThanOrEqual(5);
    expect(outlineSegments.every(segment => segment.lw === 2)).toBe(true);
  });

  it('uses the box-series outline for raw observation markers', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, boxModel({
      catAxisHidden: true,
      valAxisHidden: true,
      chartexDataPointMarkerStyle: {
        fillColors: ['FFFFFF'],
        lineColors: ['FFFFFF'],
        lineWidthEmu: 9525,
      },
      chartexBox: {
        categories: ['A'],
        series: [{
          name: 'S', color: 'FFFFFF', lineColor: '000000', lineWidthEmu: 6350,
          valuesByCategory: [[1, 2, 3]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: true,
          quartileMethod: 'inclusive',
        }],
      },
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(3);
    expect(rec.strokeDetails.some(stroke => stroke.strokeStyle.toLowerCase() === '#ffffff'))
      .toBe(false);
    expect(rec.strokeDetails.filter(stroke => stroke.strokeStyle.toLowerCase() === '#000000').length)
      .toBeGreaterThanOrEqual(3);
  });

  it('uses the same 14pt omitted-title fallback for classic and ChartEx charts', () => {
    const cases: Array<{ title: string; chart: ChartModel }> = [
      {
        title: 'classic title',
        chart: baseModel({
          chartType: 'line',
          title: 'classic title',
          titleFontSizeHpt: null,
          categories: ['A', 'B'],
          series: [series({ values: [1, 2] })],
        }),
      },
      {
        title: 'ChartEx title',
        chart: boxModel({ title: 'ChartEx title', titleFontSizeHpt: null }),
      },
    ];

    for (const { title, chart } of cases) {
      const rec = ringRecordingCtx();
      renderChart(rec.ctx, chart, RECT, 1);
      expect(rec.fontTexts.find(text => text.text === title)?.font).toContain('14px');
    }
  });

  it('uses the parser-resolved linked Chart Style title size', () => {
    // With titleFontSizeHpt=1400 (14pt) at scale 1 the title renders at 14px.
    // Capture the font active at each fillText so we can read the title's size.
    const drawn: Array<{ text: string; px: number }> = [];
    const state: Record<string, unknown> = {
      font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    };
    const px = (font: string): number => {
      const m = /(\d+(?:\.\d+)?)px/.exec(font);
      return m ? parseFloat(m[1]) : 10;
    };
    const ctx = new Proxy(state, {
      get(_t, prop: string) {
        if (prop in state && typeof state[prop] !== 'function') return state[prop];
        if (prop === 'measureText') return (t: string) => ({ width: String(t).length * px(String(state.font)) * 0.6 });
        if (prop === 'fillText') return (text: string) => { drawn.push({ text, px: px(String(state.font)) }); };
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
        return () => undefined;
      },
      set(_t, prop: string, value) { state[prop] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;

    renderChart(ctx, boxModel({ title: 'the box title', titleFontSizeHpt: 1400 }), RECT, 1);
    const titleDraw = drawn.find(d => d.text === 'the box title');
    expect(titleDraw).toBeDefined();
    // 1400 hpt → 14pt → 14px at scale 1.
    expect(titleDraw?.px).toBeCloseTo(14, 5);
  });
});

// CH15 — chartEx sunburst (MS 2014 chartex ext). Verify the hierarchy folds
// into concentric rings, each branch's sub-tree shares its accent color, and
// angular spans are size-proportional.
describe('CH15 — chartEx sunburst', () => {
  // Two branches, each with two stems, each stem with one leaf. Branch A is
  // twice the total of Branch B (so it must sweep twice the angle).
  function sunburstModel(over: Partial<ChartModel> = {}): ChartModel {
    return baseModel({
      chartType: 'sunburst',
      title: 'sun',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexSunburst: {
        rows: [
          { path: ['Branch A', 'Stem 1', 'Leaf 1'], size: 30 },
          { path: ['Branch A', 'Stem 2', 'Leaf 2'], size: 30 },
          { path: ['Branch B', 'Stem 3', 'Leaf 3'], size: 15 },
          { path: ['Branch B', 'Stem 4', 'Leaf 4'], size: 15 },
        ],
      },
      ...over,
    });
  }

  it('preserves rich hierarchy-label runs through the shared ChartEx resolver', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      series: [series({
        values: [1],
        dataLabelOverrides: [{
          idx: 0, text: 'HIDDENVISIBLE',
          richRuns: [
            { text: 'HIDDEN', colorPaintAuthored: true, colorHidden: true },
            { text: 'VISIBLE' },
          ],
        }],
      })],
      chartStyleRoles: { dataLabel: { fontColor: '008000' } },
    }), RECT, 1);
    expect(rec.fontTexts.some(text => text.text === 'HIDDEN')).toBe(false);
    expect(rec.fontTexts).toContainEqual(expect.objectContaining({ text: 'VISIBLE', fill: '#008000' }));
  });

  it('rejects oversized hierarchy rows and paths before allocating the tree', () => {
    for (const rows of [
      Array.from({ length: 10_001 }, (_, index) => ({ path: [`N${index}`], size: 1 })),
      [{ path: Array.from({ length: 10_001 }, (_, index) => `D${index}`), size: 1 }],
      [{ path: Array.from({ length: 513 }, (_, index) => `D${index}`), size: 1 }],
    ]) {
      const rec = ringRecordingCtx();
      expect(() => renderChart(rec.ctx, sunburstModel({ chartexSunburst: { rows } }), RECT, 1))
        .not.toThrow();
      expect(rec.fontTexts.map(text => text.text)).toContain('(too many data points)');
      expect(rec.arcs).toEqual([]);
    }
  });

  it('handles the bounded wide-tree limit without argument or recursion overflow', () => {
    const rec = ringRecordingCtx();
    const rows = Array.from({ length: 10_000 }, (_, index) => ({ path: [`N${index}`], size: 1 }));
    expect(() => renderChart(rec.ctx, sunburstModel({ chartexSunburst: { rows } }), RECT, 1))
      .not.toThrow();
    expect(rec.fontTexts.map(text => text.text)).not.toContain('(too many data points)');
    expect(rec.arcs.length).toBeGreaterThan(0);
  });

  it('charges label-shape paint by interned hierarchy node rather than source row', () => {
    const gradient = {
      fillType: 'gradient' as const,
      gradType: 'linear' as const,
      angle: 0,
      stops: Array.from({ length: 4096 }, (_, index) => ({
        position: index / 4095,
        color: '112233',
      })),
    };
    const build = (depth: number): ChartModel => sunburstModel({
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          labelBox: { fillPaint: gradient },
        },
      })],
      chartexSunburst: {
        rows: [{
          path: Array.from({ length: depth }, (_, index) => `D${index}`),
          size: 1,
        }],
      },
    });
    expect(chartExHierarchyLabelPaintWorkCount(build(1))).toBe(4096);
    expect(chartExHierarchyLabelPaintWorkCount(build(256))).toBe(1_048_576);
    expect(chartExHierarchyLabelPaintWorkCount(build(257))).toBe(1_048_577);
    const sparse = build(1);
    sparse.chartexSunburst = {
      rows: [
        ...Array.from({ length: 256 }, (_, index) => ({ path: [`Z${index}`], size: 0 })),
        { path: ['Visible'], size: 1 },
      ],
    };
    expect(chartExHierarchyLabelPaintWorkCount(sparse)).toBe(4096);
  });

  it.each([
    { x: 0, y: 0, w: 640, h: 360 },
    { x: 0, y: 0, w: 360, h: 640 },
    { x: 0, y: 0, w: 900, h: 240 },
  ])('keeps the Office-observed automatic hole ratio local to sunburst (%o)', (rect) => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), rect, 1);
    const radii = rec.arcs.map(arc => arc.r).filter(radius => radius > 0);
    expect(Math.min(...radii) / Math.max(...radii)).toBeCloseTo(0.18, 5);
  });

  it('draws three concentric rings (Branch / Stem / Leaf) with distinct radii', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // Each ring segment emits an outer + inner arc; across all segments the
    // distinct radii cluster into 3 outer + 3 inner boundaries → at least 3
    // distinct radius bands (inner hole excluded).
    const radii = [...new Set(rec.arcs.map(a => Math.round(a.r)))].sort((a, b) => a - b);
    // 4 radius boundaries: hole, branch/stem, stem/leaf, outer.
    expect(radii.length).toBeGreaterThanOrEqual(4);
  });

  it('colors every node in a branch with that branch\'s accent (branch A=accent1, B=accent2)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // Segment fills (excluding the white label fills). Branch A subtree (root +
    // 2 stems + 2 leaves = 5 nodes) all accent1; Branch B (5 nodes) all accent2.
    const segFills = rec.fills.filter(f => f !== '#ffffff' && f !== '#000');
    const a1 = segFills.filter(f => f === '#5B9BD5').length;
    const a2 = segFills.filter(f => f === '#ED7D31').length;
    expect(a1).toBe(5);
    expect(a2).toBe(5);
  });

  it('draws white segment labels for the branch/stem/leaf names', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      series: [series({
        values: [],
        dataLabelOverrides: [{ idx: 0, text: 'Branch A', fontFace: 'Point Sunburst Face' }],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          fontFace: 'Sunburst Face',
        },
      })],
    }), RECT, 1);
    const whiteLabels = rec.fontTexts.filter(t => t.fill === '#ffffff').map(t => t.text);
    // Labels are word-wrapped, so assert on the first word of each name.
    const joined = whiteLabels.join('').replace(/\s/g, '');
    expect(joined).toContain('Branch');
    expect(joined).toContain('Stem');
    expect(joined).toContain('Leaf');
    expect(rec.fontTexts.some(text => text.font.includes('"Sunburst Face"'))).toBe(true);
    expect(rec.fontTexts.some(text => text.font.includes('"Point Sunburst Face"'))).toBe(true);
  });

  it('does not invent ring labels when the ChartEx series omits dataLabels', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    expect(rec.fontTexts.map(text => text.text)).not.toEqual(
      expect.arrayContaining(['Branch', 'Stem', 'Leaf']),
    );
  });

  it('applies the authored ChartEx series outline ahead of the linked Chart Style', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      chartexDataPointStyle: { lineHidden: true },
      series: [series({
        values: [],
        lineColor: 'FFFFFF',
        lineWidthEmu: 12700,
        lineHidden: false,
      })],
    }), RECT, 4 / 3);

    expect(rec.strokes.length).toBeGreaterThan(0);
    expect(rec.strokes.every(stroke =>
      stroke.strokeStyle === '#FFFFFF' && Math.abs(stroke.lineWidth - 4 / 3) < 1e-6
    )).toBe(true);
  });

  it('honors explicit ChartEx series noFill over a visible linked outline', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      chartexDataPointStyle: { lineColors: ['FFFFFF'], lineWidthEmu: 12700 },
      series: [series({ values: [], lineHidden: true })],
    }), RECT, 4 / 3);

    expect(rec.strokes).toHaveLength(0);
  });

  it('orients category labels radially instead of tangentially', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      series: [series({
        values: [],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
        },
      })],
    }), RECT, 1);
    // Branch A owns 2/3 of the circle. Starting at -90°, its midpoint is 30°.
    // Radial text rotates by that midpoint; the former tangential layout used
    // 120° (midpoint + 90°).
    expect(rec.rotates[0]).toBeCloseTo(Math.PI / 6, 5);
  });

  it('bounds a long authored sunburst label and preserves whitespace', () => {
    const rec = ringRecordingCtx();
    const authored = `A  ${'x'.repeat(5000)}`;
    renderChart(rec.ctx, sunburstModel({
      series: [series({
        values: [],
        dataLabelOverrides: [{ idx: 0, text: authored, fontColor: '123456', position: 'ctr' }],
      })],
    }), RECT, 1);

    const labels = rec.fontTexts.filter(text => text.fill === '#123456');
    expect(labels.length).toBeGreaterThan(0);
    expect(labels.length).toBeLessThanOrEqual(4);
    expect(labels.map(text => text.text).join('')).toContain('A  ');
    expect(labels.map(text => text.text)).not.toContain(authored);
  });

  it('applies authored sunburst manual layout in chart coordinates', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, sunburstModel({
      series: [series({
        values: [],
        dataLabelOverrides: [{
          idx: 0,
          text: 'manual sunburst',
          fontColor: '123456',
          manualLayout: { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5, w: 0.2, h: 0.1 },
        }],
      })],
    }), RECT, 1);

    expect(rec.texts.find(text => text.text === 'manual sunburst')).toMatchObject({ x: 384, y: 198 });
  });

  it('sweeps each branch proportional to its aggregated size (Branch A twice Branch B)', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel(), RECT, 1);
    // The innermost ring (smallest non-hole outer radius) carries the two branch
    // segments. Each segment's outer arc sweep = a1 − a0. Branch A (size 60) must
    // sweep ~2× Branch B (size 30).
    const innerOuterR = [...new Set(rec.arcs.map(a => Math.round(a.r)))].sort((a, b) => a - b)[1];
    const branchArcs = rec.arcs.filter(a => Math.round(a.r) === innerOuterR && !a.ccw);
    const sweeps = branchArcs.map(a => Math.abs(a.a1 - a.a0)).sort((x, y) => y - x);
    expect(sweeps.length).toBeGreaterThanOrEqual(2);
    expect(sweeps[0] / sweeps[1]).toBeCloseTo(2, 1);
  });

  it('ignores non-positive and non-finite hierarchy weights without changing source order', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      chartexSunburst: {
        rows: [
          { path: ['First', 'Positive'], size: 10 },
          { path: ['First', 'Negative'], size: -100 },
          { path: ['First', 'NaN'], size: Number.NaN },
          { path: ['Second', 'Positive'], size: 10 },
          { path: ['Second', 'Infinity'], size: Number.POSITIVE_INFINITY },
          { path: ['Third', 'Zero'], size: 0 },
        ],
      },
    }), RECT, 1);

    const innerOuterR = [...new Set(rec.arcs.map(arc => Math.round(arc.r)))]
      .sort((a, b) => a - b)[1];
    const branchArcs = rec.arcs.filter(arc => Math.round(arc.r) === innerOuterR && !arc.ccw);
    const sweeps = branchArcs.map(arc => arc.a1 - arc.a0);
    expect(sweeps).toHaveLength(2);
    expect(sweeps[0]).toBeCloseTo(Math.PI, 5);
    expect(sweeps[1]).toBeCloseTo(Math.PI, 5);
    expect(branchArcs[0].a0).toBeCloseTo(-Math.PI / 2, 5);
    expect(branchArcs[1].a0).toBeCloseTo(branchArcs[0].a1, 5);
  });

  it('keeps proportional finite angles when positive hierarchy sums exceed Number.MAX_VALUE', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      chartexSunburst: {
        rows: [
          { path: ['First', 'A'], size: Number.MAX_VALUE },
          { path: ['First', 'B'], size: Number.MAX_VALUE },
          { path: ['Second', 'C'], size: Number.MAX_VALUE },
        ],
      },
    }), RECT, 1);

    const innerOuterR = [...new Set(rec.arcs.map(arc => Math.round(arc.r)))]
      .sort((a, b) => a - b)[1];
    const branchArcs = rec.arcs.filter(arc => Math.round(arc.r) === innerOuterR && !arc.ccw);
    expect(branchArcs).toHaveLength(2);
    expect(branchArcs.every(arc => Number.isFinite(arc.a0) && Number.isFinite(arc.a1))).toBe(true);
    expect(branchArcs[0].a1 - branchArcs[0].a0).toBeCloseTo((Math.PI * 4) / 3, 5);
    expect(branchArcs[1].a1 - branchArcs[1].a0).toBeCloseTo((Math.PI * 2) / 3, 5);
  });

  it('paints no sunburst geometry when every hierarchy weight normalizes to zero', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({
      chartexSunburst: {
        rows: [
          { path: ['Zero'], size: 0 },
          { path: ['Negative'], size: -1 },
          { path: ['NaN'], size: Number.NaN },
          { path: ['Infinity'], size: Number.POSITIVE_INFINITY },
        ],
      },
    }), RECT, 1);

    expect(rec.arcs).toHaveLength(0);
    expect(rec.fills).toHaveLength(0);
  });

  it('draws an authored top legend outside the rings', () => {
    const rec = ringRecordingCtx();
    renderChart(rec.ctx, sunburstModel({ showLegend: true, legendPos: 't' }), RECT, 1);
    const legendLabels = rec.fontTexts
      .filter(text => text.fill !== '#ffffff')
      .map(text => text.text);
    expect(legendLabels).toEqual(expect.arrayContaining(['Branch A', 'Branch B']));
  });

});

// CH15 — chartEx treemap. The parser supplies the same root→leaf rows as
// sunburst; the renderer must turn them into nested, area-proportional tiles.
describe('CH15 — chartEx treemap', () => {
  function treemapModel(): ChartModel {
    return baseModel({
      chartType: 'treemap',
      title: 'regions',
      series: [series({
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
        },
      })],
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'],
      chartexTreemap: {
        parentLabelLayout: 'banner',
        rows: [
          { path: ['Americas', 'North'], size: 50 },
          { path: ['Americas', 'South'], size: 30 },
          { path: ['Asia', 'East'], size: 20 },
        ],
      },
    });
  }

  it('atomically rejects a hierarchy whose total path segments exceed the paint cap', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap = {
      ...model.chartexTreemap,
      rows: Array.from({ length: 5_001 }, (_, index) => ({
        path: [`P${index}`, `L${index}`], size: 1,
      })),
    };
    expect(() => renderChart(rec.ctx, model, RECT, 1)).not.toThrow();
    expect(rec.texts.map(text => text.text)).toContain('(too many data points)');
    expect(rec.rects).toEqual([]);
  });

  it('draws nested branch and leaf rectangles instead of the unsupported-chart placeholder', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, treemapModel(), RECT, 1);
    expect(rec.texts.map(t => t.text)).not.toContain('Chart: treemap');
    expect(rec.texts.map(t => t.text)).toEqual(expect.arrayContaining(['Americas', 'Asia', 'North', 'South', 'East']));
    // Two parent regions + three leaves. Parent banners may add a background,
    // so assert a lower bound rather than an exact implementation count.
    expect(rec.rects.length).toBeGreaterThanOrEqual(5);
  });

  it('uses one theme accent per top-level branch', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, treemapModel(), RECT, 1);
    const fills = rec.rects.map(r => r.fs.toUpperCase());
    expect(fills).toContain('#5B9BD5');
    expect(fills).toContain('#ED7D31');
  });

  it('keeps repeated terminal labels as distinct tiles and legend entries', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['Repeat'], size: 3 },
          { path: ['Repeat'], size: 2 },
          { path: ['Repeat'], size: 1 },
        ],
      },
      showLegend: true,
      legendPos: 't',
      series: [series({
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
        },
      })],
    });
    renderChart(rec.ctx, model, RECT, 1);

    const tiles = rec.rects.filter(rect => rect.w > 20 && rect.h > 20);
    expect(tiles).toHaveLength(3);
    expect(rec.texts.filter(text => text.text === 'Repeat').length).toBeGreaterThanOrEqual(6);
  });

  it('normalizes hierarchy weights before allocating proportional parent areas', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['A', 'positive'], size: 10 },
          { path: ['A', 'negative'], size: -100 },
          { path: ['A', 'invalid'], size: Number.NaN },
          { path: ['B', 'positive'], size: 10 },
          { path: ['B', 'infinite'], size: Number.POSITIVE_INFINITY },
          { path: ['C', 'zero'], size: 0 },
        ],
      },
      series: [series({})],
    });
    renderChart(rec.ctx, model, RECT, 1);

    const areaByFill = new Map<string, number>();
    for (const tile of rec.rects) {
      areaByFill.set(tile.fs, (areaByFill.get(tile.fs) ?? 0) + tile.w * tile.h);
    }
    expect([...areaByFill.keys()]).toEqual(['#5B9BD5', '#ED7D31']);
    expect(areaByFill.get('#5B9BD5')).toBeCloseTo(areaByFill.get('#ED7D31') ?? 0, 5);
  });

  it('keeps squarified tiles proportional, contained, non-overlapping, and deterministic', () => {
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['Large', '1000'], size: 1000 },
          { path: ['Small', '1'], size: 1 },
          { path: ['Zero', '0'], size: 0 },
        ],
      },
      series: [series({})],
    });
    const bounds = { x: 0, y: 0, w: 2200, h: 1200 };
    const first = recordingCtx();
    const second = recordingCtx();
    renderChart(first.ctx, model, bounds, 1);
    renderChart(second.ctx, model, bounds, 1);

    expect(first.rects).toEqual(second.rects);
    expect(first.rects).toHaveLength(2);
    const [large, small] = first.rects.sort((a, b) => b.w * b.h - a.w * a.h);
    expect((large.w * large.h) / (small.w * small.h)).toBeCloseTo(1000, 5);

    const minX = Math.min(...first.rects.map(tile => tile.x));
    const minY = Math.min(...first.rects.map(tile => tile.y));
    const maxX = Math.max(...first.rects.map(tile => tile.x + tile.w));
    const maxY = Math.max(...first.rects.map(tile => tile.y + tile.h));
    const unionArea = (maxX - minX) * (maxY - minY);
    const tileArea = first.rects.reduce((sum, tile) => sum + tile.w * tile.h, 0);
    expect(tileArea).toBeCloseTo(unionArea, 5);
    for (let i = 0; i < first.rects.length; i++) {
      const a = first.rects[i];
      expect(a.x).toBeGreaterThanOrEqual(minX);
      expect(a.y).toBeGreaterThanOrEqual(minY);
      expect(a.x + a.w).toBeLessThanOrEqual(maxX);
      expect(a.y + a.h).toBeLessThanOrEqual(maxY);
      for (let j = i + 1; j < first.rects.length; j++) {
        const b = first.rects[j];
        const overlapW = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
        const overlapH = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
        expect(overlapW * overlapH).toBeCloseTo(0, 8);
      }
    }
  });

  it('preserves a 10:1 aggregate-weight ratio between top-level parents', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['Ten', 'A'], size: 6 },
          { path: ['Ten', 'B'], size: 4 },
          { path: ['One', 'C'], size: 1 },
        ],
      },
      series: [series({})],
    }), RECT, 1);
    const area = (color: string): number => rec.rects
      .filter(tile => tile.fs === color)
      .reduce((sum, tile) => sum + tile.w * tile.h, 0);
    expect(area('#5B9BD5') / area('#ED7D31')).toBeCloseTo(10, 5);
  });

  it('keeps deep aggregate areas finite when positive sums exceed Number.MAX_VALUE', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['First', 'Deep', 'A'], size: Number.MAX_VALUE },
          { path: ['First', 'Deep', 'B'], size: Number.MAX_VALUE },
          { path: ['Second', 'C'], size: Number.MAX_VALUE },
        ],
      },
      series: [series({})],
    }), RECT, 1);

    expect(rec.rects).toHaveLength(3);
    expect(rec.rects.every(tile => [tile.x, tile.y, tile.w, tile.h].every(Number.isFinite))).toBe(true);
    const area = (color: string): number => rec.rects
      .filter(tile => tile.fs === color)
      .reduce((sum, tile) => sum + tile.w * tile.h, 0);
    expect(area('#5B9BD5') / area('#ED7D31')).toBeCloseTo(2, 5);
  });

  it('keeps an overflowing leaf display label finite', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['Finite'], size: Number.MAX_VALUE }],
      },
      series: [series({
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
        },
      })],
    }), RECT, 1);

    const leafValues = rec.texts
      .map(text => Number(text.text));
    expect(leafValues).toContain(Number.MAX_VALUE);
    expect(leafValues.filter(Number.isFinite).every(Number.isFinite)).toBe(true);
  });

  it('keeps overlapping treemap parent captions category-only when leaf labels show values', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      chartexTreemap: {
        parentLabelLayout: 'overlapping',
        rows: [
          { path: ['Civilian', 'Public services'], size: 0.62 },
          { path: ['Civilian', 'Infrastructure'], size: 0.38 },
          { path: ['Military', 'Personnel'], size: 0.55 },
          { path: ['Military', 'Equipment'], size: 0.45 },
        ],
      },
      series: [series({
        valFormatCode: '0%',
        seriesDataLabels: {
          showVal: true,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'inEnd',
          separator: '\n',
        },
      })],
    }), RECT, 1);

    const parentCaptions = rec.texts
      .filter(text => text.baseline === 'top')
      .map(text => text.text);
    expect(parentCaptions).toEqual(expect.arrayContaining(['Civilian', 'Military']));
    expect(parentCaptions.every(text => !text.includes('%') && !text.includes('\n'))).toBe(true);
    expect(rec.texts.map(text => text.text)).toEqual(
      expect.arrayContaining(['Public services', '62%']),
    );
  });

  it('keeps equal-weight top-level tiles in first-seen source order', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5', 'ED7D31', 'A5A5A5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [
          { path: ['First'], size: 1 },
          { path: ['Second'], size: 1 },
          { path: ['Third'], size: 1 },
        ],
      },
      series: [series({})],
    }), RECT, 1);

    expect(rec.rects.map(tile => tile.fs)).toEqual(['#5B9BD5', '#ED7D31', '#A5A5A5']);
  });

  it('does not paint padded parent frames for overlapping parent labels', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    // `overlapping` places the parent caption over its descendant tiles. The
    // parent is not an additional painted tile or frame: only the three leaf
    // rectangles are visible and separated by their own borders.
    expect(rec.rects).toHaveLength(3);
    expect(rec.strokeRects).toHaveLength(3);
  });

  it('keeps the exact top-level accent on every descendant data point', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.rects.slice(0, 2).map(rect => rect.fs)).toEqual(['#5B9BD5', '#5B9BD5']);
    expect(rec.rects[2].fs).toBe('#ED7D31');
  });

  it('labels top-level branches without overlapping intermediate hierarchy labels', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap = {
      parentLabelLayout: 'overlapping',
      rows: [
        { path: ['Group A', 'Subgroup A1', 'A-major'], size: 1000 },
        { path: ['Group A', 'Subgroup A1', 'A-medium'], size: 100 },
        { path: ['Group B', 'Subgroup B1', 'B-major'], size: 1000 },
      ],
    };
    renderChart(rec.ctx, model, RECT, 1);

    const labels = rec.texts.map(text => text.text);
    expect(labels).toEqual(expect.arrayContaining(['Group A', 'Group B']));
    expect(labels).not.toEqual(expect.arrayContaining(['Subgroup A1', 'Subgroup B1']));
  });

  it('separates overlapping parent captions from authored inEnd leaves in a two-level treemap', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      series: [series({
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'inEnd',
          fontSizeHpt: 900,
        },
      })],
      chartexTreemap: {
        parentLabelLayout: 'overlapping',
        rows: [
          { path: ['Female', 'California'], size: 81_400 },
          { path: ['Female', 'Indiana'], size: 62_000 },
          { path: ['Female', 'Pennsylvania'], size: 59_900 },
          { path: ['Male', 'California'], size: 97_500 },
          { path: ['Male', 'Indiana'], size: 45_200 },
          { path: ['Male', 'Pennsylvania'], size: 45_900 },
        ],
      },
    }), RECT, 1);

    const parents = rec.texts.filter(text => text.text === 'Female' || text.text === 'Male');
    const leaves = rec.texts.filter(text =>
      text.text === 'California' || text.text === 'Indiana' || text.text === 'Pennsylvania'
    );
    expect(parents).toHaveLength(2);
    expect(leaves).toHaveLength(6);
    expect(parents.every(text => text.align === 'left' && text.baseline === 'top')).toBe(true);
    expect(leaves.every(text => text.align === 'left' && text.baseline === 'bottom')).toBe(true);
    expect([...parents, ...leaves].every(text => text.font?.includes('9px '))).toBe(true);
    // The top caption band and the lowest available leaf edge are disjoint;
    // this locks the hierarchy-role split that prevents two labels occupying
    // the same lower-left anchor.
    expect(Math.max(...parents.map(text => text.y + 9)))
      .toBeLessThanOrEqual(Math.min(...leaves.map(text => text.y - 9)));
  });

  it('maps a treemap outEnd value label to the Office-observed lower-left tile anchor', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      series: [series({
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          position: 'outEnd',
          fontSizeHpt: 1000,
        },
      })],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['472'], size: 472 }],
      },
    }), RECT, 1);

    const value = rec.texts.find(text => text.text === '472');
    expect(value).toBeDefined();
    expect(value).toMatchObject({ align: 'left', baseline: 'bottom' });
    expect((value as TextCall).x).toBeLessThan(RECT.w / 4);
    expect((value as TextCall).y).toBeGreaterThan(RECT.h * 0.7);
  });

  it('keeps an authored 10pt bold treemap leaf label at the DrawingML point-to-pixel size', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'treemap',
      dataLabelFontFace: 'Wrong Chart Face',
      series: [series({
        seriesDataLabels: {
          showVal: true,
          showCatName: false,
          showSerName: false,
          showPercent: false,
          position: 'outEnd',
          fontSizeHpt: 1000,
          fontBold: true,
          fontFace: 'Aptos Narrow',
        },
      })],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['42'], size: 42 }],
      },
    }), RECT, 4 / 3);

    const label = rec.texts.find(text => text.text === '42');
    expect(label).toBeDefined();
    expect(label?.font).toContain('bold ');
    expect(label?.font).toContain('"Aptos Narrow", Calibri, Arial, sans-serif');
    const px = Number.parseFloat(/(\d+(?:\.\d+)?)px/.exec(label?.font ?? '')?.[1] ?? '0');
    expect(px).toBeCloseTo(1000 / 100 * 4 / 3, 6);
  });

  it('uses the ChartEx data-point outline on the exact tile boundary', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineColors: ['FFFFFF'], lineWidthEmu: 19050 };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.strokeRects).toHaveLength(rec.rects.length);
    rec.strokeRects.forEach((stroke, index) => {
      expect(stroke).toMatchObject({
        x: rec.rects[index].x,
        y: rec.rects[index].y,
        w: rec.rects[index].w,
        h: rec.rects[index].h,
        ss: '#FFFFFF',
        lw: 1.5,
      });
    });
  });

  it('uses a one-CSS-pixel chart-background separator when no line is authored', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartBg = '112233';
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = null;
    model.series[0].lineColor = null;
    model.series[0].lineWidthEmu = null;
    model.series[0].lineHidden = null;
    renderChart(rec.ctx, model, RECT, 1);

    const tiles = rec.rects.filter(rect => rect.fs !== '#112233');
    expect(rec.strokeRects).toHaveLength(tiles.length);
    expect(rec.strokeRects.every(stroke => stroke.ss === '#112233' && stroke.lw === 1)).toBe(true);
  });

  it('uses the automatic separator when the linked Chart Style line is NoStyle', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartBg = '112233';
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineHidden: true, lineNoStyle: true };
    renderChart(rec.ctx, model, RECT, 1);

    const tiles = rec.rects.filter(rect => rect.fs !== '#112233');
    expect(rec.strokeRects).toHaveLength(tiles.length);
    expect(rec.strokeRects.every(stroke => stroke.ss === '#112233' && stroke.lw === 1)).toBe(true);
  });

  it('keeps a direct series noFill authoritative over the automatic separator', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineHidden: true, lineNoStyle: true };
    model.series[0].chartexStyle = { lineHidden: true };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.strokeRects).toHaveLength(0);
  });

  it('uses the per-accent ChartEx outline after phClr substitution', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineColors: ['112233', '445566'] };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.strokeRects.map(stroke => stroke.ss)).toEqual([
      '#112233',
      '#112233',
      '#445566',
    ]);
  });

  it('uses the shared DrawingML pattern fill for treemap data points', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexDataPointStyle = {
      fillPaints: [{ fillType: 'pattern', fg: '777777', bg: 'FFFFFF', preset: 'pct30' }],
    };
    renderChart(rec.ctx, model, RECT, 1);

    // The headless recording context has no auxiliary bitmap canvas, so the
    // shared pattern resolver falls back to its authored foreground color.
    expect(rec.rects.every(rect => rect.fs === 'rgba(119,119,119,1)')).toBe(true);
  });

  it('suppresses treemap outlines for an explicit ChartEx data-point noFill', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    model.chartexDataPointStyle = { lineHidden: true };
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.rects).toHaveLength(3);
    expect(rec.strokeRects).toHaveLength(0);
  });

  it('bounds an over-wide inEnd leaf label to the shared four-line limit', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'overlapping',
        rows: [
          { path: ['Group A', 'A-major'], size: 1000 },
          { path: ['Group A', 'A-medium'], size: 100 },
        ],
      },
      series: [series({
        values: [1000, 100],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'inEnd',
          fontSizeHpt: 1000,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 180, h: 180 }, 1);

    const narrow = [...rec.strokeRects].sort((a, b) => a.w - b.w)[0];
    const narrowLabels = rec.texts
      .filter(text => text.baseline === 'bottom' && text.x >= narrow.x && text.x <= narrow.x + narrow.w)
      .sort((a, b) => a.y - b.y);
    const narrowText = narrowLabels
      .map(text => text.text)
      .join('');
    expect(narrowText).toBe('A-m…');
    expect(narrowLabels).toHaveLength(4);
  });

  it('clips centered leaf labels and limits them to the tile height', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['Group', 'A very long centered leaf label'], size: 1 }],
      },
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'ctr',
          fontSizeHpt: 1600,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 90, h: 65 }, 1);

    const tile = rec.strokeRects[0];
    const labels = rec.texts.filter(text => text.baseline === 'middle');
    const maxLines = Math.floor((tile.h - 6) / (16 * 1.1));
    expect(labels.length).toBeLessThanOrEqual(maxLines);
    expect(labels.every(text => text.y >= tile.y && text.y <= tile.y + tile.h)).toBe(true);
    expect(rec.clips).toEqual(expect.arrayContaining([
      expect.objectContaining({ x: tile.x, y: tile.y, w: tile.w, h: tile.h }),
    ]));
  });

  it('keeps omitted-position leaf labels inside a font-relative 0.5em inset', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'none',
        rows: [{ path: ['Group', 'ABCDEFGHIJKL'], size: 1 }],
      },
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          fontSizeHpt: 2000,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 90, h: 90 }, 1);

    const tile = rec.strokeRects[0];
    const labels = rec.texts.filter(text => text.baseline === 'middle');
    expect(labels.length).toBeGreaterThan(0);
    for (const label of labels) {
      const fontPx = Number.parseFloat(/(\d+(?:\.\d+)?)px/.exec(label.font ?? '')?.[1] ?? '0');
      expect(label.x - (label.width ?? 0) / 2).toBeGreaterThanOrEqual(tile.x + fontPx * 0.5 - 1e-6);
      expect(label.x + (label.width ?? 0) / 2).toBeLessThanOrEqual(tile.x + tile.w - fontPx * 0.5 + 1e-6);
      expect(label.y).toBeGreaterThanOrEqual(tile.y + fontPx * 0.5 - 1e-6);
      expect(label.y).toBeLessThanOrEqual(tile.y + tile.h - fontPx * 0.5 + 1e-6);
    }
  });

  it('clips boundary-centered tile strokes to the plot rectangle', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.chartexTreemap!.parentLabelLayout = 'overlapping';
    renderChart(rec.ctx, model, RECT, 1);

    const minX = Math.min(...rec.strokeRects.map(rect => rect.x));
    const minY = Math.min(...rec.strokeRects.map(rect => rect.y));
    const maxX = Math.max(...rec.strokeRects.map(rect => rect.x + rect.w));
    const maxY = Math.max(...rec.strokeRects.map(rect => rect.y + rect.h));
    const plotClip = rec.clips.find(clip => Math.abs(clip.x - minX) < 0.001 && Math.abs(clip.y - minY) < 0.001);
    expect(plotClip).toBeDefined();
    expect((plotClip as { w: number }).w).toBeCloseTo(maxX - minX, 5);
    expect((plotClip as { h: number }).h).toBeCloseTo(maxY - minY, 5);
  });

  it('honors ChartEx label visibility, separator, and authored inEnd placement', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series = [series({
      values: [50, 30, 20],
      seriesDataLabels: {
        showVal: true,
        showCatName: true,
        showSerName: false,
        showPercent: false,
        separator: '\n',
        position: 'inEnd',
        fontSizeHpt: 1000,
      },
    })];
    renderChart(rec.ctx, model, RECT, 1);
    const north = rec.texts.find(text => text.text === 'North');
    const fifty = rec.texts.find(text => text.text === '50');
    expect(north).toMatchObject({ align: 'left', baseline: 'bottom' });
    expect(fifty).toMatchObject({ align: 'left', baseline: 'bottom' });
    expect((fifty as TextCall).y).toBeGreaterThan((north as TextCall).y);
  });

  it('keeps centered automatic parent and leaf text inside an inBase treemap tile', () => {
    const rec = recordingCtx();
    const model = baseModel({
      chartType: 'treemap',
      chartexAccents: ['5B9BD5'],
      chartexTreemap: {
        parentLabelLayout: 'overlapping',
        rows: [{ path: ['Group', 'Leaf'], size: 1 }],
      },
      series: [series({
        values: [1],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,
          showSerName: false,
          showPercent: false,
          position: 'inBase',
          textAlign: 'ctr',
          fontSizeHpt: 1000,
        },
      })],
    });
    renderChart(rec.ctx, model, { x: 0, y: 0, w: 160, h: 120 }, 1);

    const tile = rec.strokeRects[0];
    const labels = rec.texts.filter(text => text.text === 'Group' || text.text === 'Leaf');
    expect(labels.map(text => text.text)).toEqual(expect.arrayContaining(['Group', 'Leaf']));
    for (const label of labels) {
      expect(label.align).toBe('center');
      expect(label.x - (label.width ?? 0) / 2).toBeGreaterThanOrEqual(tile.x - 1e-6);
      expect(label.x + (label.width ?? 0) / 2).toBeLessThanOrEqual(tile.x + tile.w + 1e-6);
    }
  });

  it('applies ChartEx per-label overrides by hierarchy-node preorder index', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series = [series({
      values: [50, 30, 20],
      seriesDataLabels: {
        showVal: true,
        showCatName: true,
        showSerName: false,
        showPercent: false,
        separator: '\n',
        position: 'inEnd',
        fontColor: 'FFFFFF',
      },
      // preorder: Americas=0, North=1, South=2, Asia=3, East=4
      dataLabelOverrides: [{ idx: 4, text: 'Custom East\n20', fontColor: '222222' }],
    })];
    renderChart(rec.ctx, model, RECT, 1);
    const custom = rec.texts.filter(text =>
      (text.text === 'Custom East' || text.text === '20') && text.fillStyle === '#222222'
    );
    expect(custom.map(text => text.text)).toEqual(expect.arrayContaining(['Custom East', '20']));
    expect(rec.texts.map(text => text.text)).not.toContain('East');
  });

  it('applies authored treemap manual layout in chart coordinates', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series[0].dataLabelOverrides = [{
      idx: 1,
      text: 'manual tile',
      fontColor: '123456',
      manualLayout: { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5, w: 0.2, h: 0.1 },
    }];
    renderChart(rec.ctx, model, RECT, 1);

    expect(rec.texts.find(text => text.text === 'manual tile')).toMatchObject({ x: 384, y: 198 });
  });

  it('honors dataLabelHidden for treemap parent nodes', () => {
    const rec = recordingCtx();
    const model = treemapModel();
    model.series[0].dataLabelOverrides = [{ idx: 0, text: '', deleted: true }];
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.texts.map(text => text.text)).not.toContain('Americas');
    expect(rec.texts.map(text => text.text)).toContain('Asia');
  });
});

// ─── canvas state leak (#766) ───────────────────────────────────────────────
//
// renderChart() previously had no top-level save/restore: per-family
// renderers (pie labels, "(no data)"/default-case text, etc.) set
// textAlign='center' / textBaseline='middle' and never restored them, so the
// mutated state leaked into whatever the caller drew next on the same ctx.
// docx/pptx call renderChart() bare (no wrapping save/restore of their own),
// so a chart followed by more text on the same page/slide would render that
// text center-aligned and vertically mis-baselined. xlsx happened to be
// immune only because its call site already wraps renderChart() in its own
// save/clip/restore.
//
// Unlike the Proxy-based recordingCtx() above (which no-ops save/restore),
// this mock implements a real state stack so the fix is actually exercised.
function stackfulMockCtx(): { ctx: CanvasRenderingContext2D; texts: TextCall[] } {
  const texts: TextCall[] = [];
  const defaults = {
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    lineCap: 'butt',
    lineJoin: 'miter',
    globalAlpha: 1,
  };
  let state: Record<string, unknown> = { ...defaults };
  const stack: Record<string, unknown>[] = [];
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'save':
          return () => stack.push({ ...state });
        case 'restore':
          return () => { const s = stack.pop(); if (s) state = s; };
        case 'measureText':
          return (t: string) => ({ width: String(t).length * 6 });
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({ text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline) });
        case 'createLinearGradient':
        case 'createRadialGradient':
          return () => ({ addColorStop() {} });
        default:
          return () => undefined;
      }
    },
    set(_t, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, texts };
}

describe('canvas state leak (#766) — renderChart restores ctx state', () => {
  it('restores textAlign/textBaseline/font/fillStyle after a pie chart (which sets them for its outer-ring labels)', () => {
    const { ctx } = stackfulMockCtx();
    // Snapshot the exact props renderChart is known to mutate.
    const before = {
      textAlign: ctx.textAlign, textBaseline: ctx.textBaseline,
      font: ctx.font, fillStyle: ctx.fillStyle,
    };
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
    expect(ctx.font).toBe(before.font);
    expect(ctx.fillStyle).toBe(before.fillStyle);
  });

  it('restores state via the "(no data)" early-return path (empty series)', () => {
    const { ctx } = stackfulMockCtx();
    const before = { textAlign: ctx.textAlign, textBaseline: ctx.textBaseline };
    renderChart(ctx, baseModel({ chartType: 'pie', series: [] }), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
  });

  it('restores state via the unknown-chart-type default-case path', () => {
    const { ctx } = stackfulMockCtx();
    const before = { textAlign: ctx.textAlign, textBaseline: ctx.textBaseline };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any -- deliberately invalid chartType to hit the default branch
    renderChart(ctx, baseModel({ chartType: 'bogus' as any, series: [series({ values: [1] })] }), RECT, 1);
    expect(ctx.textAlign).toBe(before.textAlign);
    expect(ctx.textBaseline).toBe(before.textBaseline);
  });

  it('a fillText drawn immediately after renderChart is not center-aligned (regression for the leaked pie-label state)', () => {
    const { ctx, texts } = stackfulMockCtx();
    renderChart(ctx, pieCalloutModel(), RECT, 1);
    // Simulate the caller (docx/pptx) drawing more text right after the chart,
    // exactly as it would when a chart shares a page/slide with other content.
    ctx.fillText('Other countries', 10, 500);
    const after = texts[texts.length - 1];
    expect(after.text).toBe('Other countries');
    expect(after.align).toBe('start');
    expect(after.baseline).toBe('alphabetic');
  });
});

describe('CH — combo bar+line primary value axis spans BOTH the bars and the primary-axis line (§21.2.2.16 / §21.2.2.76)', () => {
  // A stacked-column + line combo (bar + line series sharing ONE `<c:valAx>`,
  // e.g. xlsx sample-9 "MONTHLY OVERVIEW"). The bars stack to a per-category
  // maximum well BELOW the line's tallest point. Excel scales the shared
  // primary value axis to encompass every series plotted on it, regardless of
  // chart type — so the axis top must cover the line, not just the bar stack.
  //
  // Recreates sample-9's data: 3 stacked bar series (max category sum 150) plus
  // one primary-axis line "Amount Spent" whose tallest point is 180. Excel draws
  // the axis $0..$200; before the fix the renderer sized the axis to the bar sum
  // alone (150 → 160) and the 180 line point overshot the top gridline into the
  // chart title.
  const numericValLabels = (rec: Recorded): number[] =>
    rec.texts
      .map(t => Number(String(t.text).replace(/[^0-9.\-]/g, '')))
      .filter(v => Number.isFinite(v));

  it('the primary-axis line pushes the axis maximum above the stacked-bar sum', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
      series: [
        series({ name: 'Birthday',  seriesType: 'bar',  values: [20, 0, 0, 0, 0, 50, 0] }),
        series({ name: 'Holiday',   seriesType: 'bar',  values: [0, 0, 0, 0, 0, 0, 50] }),
        series({ name: 'Other',     seriesType: 'bar',  values: [0, 0, 0, 20, 0, 100, 0] }),
        // Primary-axis line: NOT on a secondary axis. Tallest point 180 > the
        // stacked bar's per-category max sum of 150.
        series({ name: 'Amount Spent', seriesType: 'line', values: [30, 0, 0, 20, 0, 180, 70] }),
      ],
    }), RECT, 1);
    const nums = numericValLabels(rec);
    const axisMax = Math.max(...nums);
    // Excel draws $0..$200 for this data. The axis top must at minimum cover the
    // line's 180 (the bug capped it at 160, hiding the tallest line point).
    expect(axisMax).toBeGreaterThanOrEqual(180);
    // And it must be the observed automatic bound above 180 — 200 — not an
    // over-scaled value.
    expect(axisMax).toBe(200);
  });

  it('WITHOUT the line, the same bars alone scale the axis to just cover 150 (isolates the line contribution)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
      series: [
        series({ name: 'Birthday', seriesType: 'bar', values: [20, 0, 0, 0, 0, 50, 0] }),
        series({ name: 'Holiday',  seriesType: 'bar', values: [0, 0, 0, 0, 0, 0, 50] }),
        series({ name: 'Other',    seriesType: 'bar', values: [0, 0, 0, 20, 0, 100, 0] }),
      ],
    }), RECT, 1);
    const axisMax = Math.max(...numericValLabels(rec));
    // Bars alone: max category sum 150 → the measured ten-interval policy
    // chooses 20-unit steps and an axis top of 160.
    expect(axisMax).toBe(160);
  });

  it('a negative primary-axis line pulls the axis minimum below the bars', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'stackedBar',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Bar',  seriesType: 'bar',  values: [40, 40] }),
        // Positive-only bars would anchor the axis at 0; the line dips to -30.
        series({ name: 'Line', seriesType: 'line', values: [10, -30] }),
      ],
    }), RECT, 1);
    const nums = numericValLabels(rec);
    // Some negative tick label appears (axis min < 0), covering the -30 line point.
    expect(nums.some(v => v <= -30)).toBe(true);
  });

  it('a SECONDARY-axis line does NOT inflate the primary axis (guards over-scaling)', () => {
    // When a line rides its own right-hand `<c:valAx>` (secondaryValAxis +
    // useSecondaryAxis), its large values live on the secondary scale and must
    // NOT stretch the primary (bar) axis — the two axes are independent. We prove
    // this by the invariant: the primary bar geometry is byte-identical whether
    // the secondary line is present or absent. If the 950 line leaked onto the
    // primary axis, the primary scale would jump ~24× and the bars would shrink.
    const bars = (secondaryLine: boolean): RectCall[] => {
      const rec = recordingCtx();
      const s: ChartSeries[] = [series({ name: 'Bar', seriesType: 'bar', values: [40, 40] })];
      if (secondaryLine) {
        s.push(series({ name: 'Line', seriesType: 'line', values: [900, 950], useSecondaryAxis: true }));
      }
      renderChart(rec.ctx, baseModel({
        chartType: 'stackedBar',
        categories: ['A', 'B'],
        series: s,
        secondaryValAxis: secondaryLine ? {
          min: null, max: null, title: null, hidden: false,
          lineHidden: false, majorTickMark: 'out', majorUnit: null,
        } : null,
      }), RECT, 1);
      return rec.rects;
    };
    const withLine = bars(true);
    const withoutLine = bars(false);
    // Same two bars, same heights — the secondary line did not touch the primary.
    expect(withLine.length).toBe(2);
    expect(withoutLine.length).toBe(2);
    expect(withLine[0].h).toBeCloseTo(withoutLine[0].h, 4);
    expect(withLine[1].h).toBeCloseTo(withoutLine[1].h, 4);
  });
});

describe('CH — combo chart legends reflect each series chart group', () => {
  it('draws the line-series legend key as a line inside a bar+line chart', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [
        series({ name: 'Amount', seriesType: 'bar', values: [20, 30] }),
        series({ name: 'Time', seriesType: 'line', values: [5, 10] }),
      ],
      showLegend: true,
      legendPos: 'b',
    }), RECT, 1);

    const lineLabel = rec.texts.find(t => t.text === 'Time');
    expect(lineLabel).toBeDefined();
    const lineKey = rec.segments.find(seg =>
      seg.length === 2 &&
      Math.abs(seg[0].y - (lineLabel as TextCall).y) < 0.01 &&
      Math.abs(seg[1].y - (lineLabel as TextCall).y) < 0.01 &&
      seg[1].x - seg[0].x > 10 &&
      seg[1].x < (lineLabel as TextCall).x,
    );
    expect(lineKey).toBeDefined();
  });

  it('keeps a line visible on both sides of a circular line-series legend marker', () => {
    const rec = markerRecordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [series({
        name: 'Outstanding', values: [20, 30],
        markerSymbol: 'circle', markerSize: 7, markerFill: 'FFFFFF', markerLine: '1696D2',
      })],
      showLegend: true,
      legendPos: 'b',
    }), RECT, 1);

    const label = rec.texts.find(text => text.text === 'Outstanding');
    const keyLine = rec.segments.find(segment =>
      segment.length === 2 &&
      Math.abs(segment[0].y - (label as TextCall).y) < 0.01 &&
      Math.abs(segment[1].y - (label as TextCall).y) < 0.01 &&
      segment[1].x < (label as TextCall).x,
    );
    const keyMarker = rec.arcs.find(arc =>
      Math.abs(arc.y - (label as TextCall).y) < 0.01 &&
      arc.x < (label as TextCall).x,
    );
    expect(keyLine).toBeDefined();
    expect(keyMarker).toBeDefined();
    expect((keyLine as Array<{ x: number; y: number }>)[1].x - (keyLine as Array<{ x: number; y: number }>)[0].x)
      .toBeGreaterThan((keyMarker as ArcCall).r * 2);
  });
});
