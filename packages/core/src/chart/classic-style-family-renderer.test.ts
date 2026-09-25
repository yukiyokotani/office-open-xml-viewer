import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartSeries } from '../types/chart.js';
import { renderChart } from './renderer.js';

const RECT = { x: 0, y: 0, w: 480, h: 300 };

function series(over: Partial<ChartSeries> = {}): ChartSeries {
  return { name: '', color: null, values: [1, 2, 3], ...over };
}

function model(over: Partial<ChartModel>): ChartModel {
  const base: ChartModel = {
    chartType: 'line',
    title: null,
    categories: ['A', 'B', 'C'],
    series: [series()],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    showLegend: false,
    legendPos: null,
    catAxisHidden: true,
    valAxisHidden: true,
    catAxisLineHidden: true,
    valAxisLineHidden: true,
    valAxisMajorGridlines: false,
    plotAreaBg: null,
    chartBg: null,
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
  };
  return { ...base, ...over };
}

function recordingContext(): {
  ctx: CanvasRenderingContext2D;
  fills: string[];
  rectFills: string[];
  strokes: string[];
  texts: string[];
} {
  const fills: string[] = [];
  const rectFills: string[] = [];
  const strokes: string[] = [];
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    lineCap: 'butt', lineJoin: 'miter', font: '10px sans-serif',
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    globalCompositeOperation: 'source-over',
  };
  const stack: Array<Record<string, unknown>> = [];
  const ctx = new Proxy(state, {
    get(target, prop: string) {
      if (prop in target) return target[prop];
      if (prop === 'save') return () => stack.push({ ...target });
      if (prop === 'restore') return () => Object.assign(target, stack.pop() ?? {});
      if (prop === 'fill') return () => fills.push(String(target.fillStyle));
      if (prop === 'fillRect') return () => rectFills.push(String(target.fillStyle));
      if (prop === 'stroke' || prop === 'strokeRect') {
        return () => strokes.push(String(target.strokeStyle));
      }
      if (prop === 'fillText') return (value: string) => texts.push(value);
      if (prop === 'measureText') return (text: string) => ({ width: text.length * 6 });
      if (prop === 'getLineDash') return () => [];
      if (prop === 'setLineDash') return () => {};
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      return () => {};
    },
    set(target, prop: string, value) {
      target[prop] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fills, rectFills, strokes, texts };
}

describe('classic chart style family wiring', () => {
  it('indexes a lone varyColors bar by point and honors permitted direct noFill', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'clusteredBar',
      varyColors: true,
      series: [series({
        chartexFormatIdx: 8,
        dataPointOverrides: [{ idx: 1, fillHidden: true }],
      })],
      chartStyleRoles: {
        dataPoint: {
          fillColors: ['888888', 'AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [8, 0, 1, 2],
          allowNoFillOverride: true,
        },
      },
    }), RECT, 1);

    expect(rec.rectFills).toContain('#AA0000');
    expect(rec.rectFills).toContain('#0000AA');
    expect(rec.rectFills).not.toContain('#00AA00');
  });

  it('uses a sparse source series index for a lone line', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      series: [series({ chartexFormatIdx: 8, showMarker: false })],
      chartStyleRoles: {
        dataPointLine: {
          lineColors: ['8899AA', '001100', '002200', '003300'],
          lineFormattingIndices: [8, 0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.strokes).toContain('#8899AA');
    expect(rec.strokes).not.toContain('#001100');
  });

  it('keeps a lone varyColors bar legend on the point-index palette', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'clusteredBar',
      varyColors: true,
      showLegend: true,
      legendPos: 'r',
      series: [series({ chartexFormatIdx: 8 })],
      chartStyleRoles: {
        dataPoint: {
          fillColors: ['888888', 'AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [8, 0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.rectFills.filter(color => color === '#AA0000')).toHaveLength(2);
    expect(rec.rectFills.filter(color => color === '#00AA00')).toHaveLength(2);
    expect(rec.rectFills.filter(color => color === '#0000AA')).toHaveLength(2);
    expect(rec.rectFills).not.toContain('#888888');
  });

  it('replays a point-index style palette in every multi-series doughnut ring', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'doughnut',
      varyColors: true,
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Outer', values: [1, 1, 1], chartexFormatIdx: 8 }),
        series({ name: 'Inner', values: [1, 1, 1], chartexFormatIdx: 9 }),
      ],
      chartStyleRoles: {
        dataPoint: {
          fillColors: ['AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#AA0000')).toHaveLength(2);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(2);
    expect(rec.fills.filter(color => color === '#0000AA')).toHaveLength(2);
  });

  it('uses a series-index style palette and series legend for varyColors=false doughnuts', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'doughnut',
      varyColors: false,
      showLegend: true,
      legendPos: 'r',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Outer', values: [1, 1, 1], chartexFormatIdx: 8 }),
        series({ name: 'Inner', values: [1, 1, 1], chartexFormatIdx: 9 }),
      ],
      chartStyleRoles: {
        dataPoint: {
          fillColors: ['AA0000', '00AA00'],
          fillFormattingIndices: [8, 9],
        },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#AA0000')).toHaveLength(3);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(3);
    expect(rec.texts).toContain('Outer');
    expect(rec.texts).toContain('Inner');
    expect(rec.texts).not.toContain('A');
  });

  it.each([
    ['line', {}],
    ['scatter', { scatterStyle: 'line' }],
    ['radar', {}],
    ['stock', { stockHiLowLines: false }],
  ] as const)('uses dataPointLine for the %s family', (chartType, extra) => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType,
      ...extra,
      series: chartType === 'stock'
        ? [series(), series({ values: [3, 4, 5] }), series({ values: [2, 3, 4] })]
        : [series({ showMarker: false })],
      chartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);

    expect(rec.strokes).toContain('#A1B2C3');
  });

  it('keeps positive direct line paint above linked style and gates style-layer noFill by its modifier', () => {
    const direct = recordingContext();
    renderChart(direct.ctx, model({
      series: [series({ lineColor: 'CC3300', showMarker: false })],
      chartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(direct.strokes).toContain('#CC3300');
    expect(direct.strokes).not.toContain('#A1B2C3');

    const hidden = recordingContext();
    renderChart(hidden.ctx, model({
      series: [series({ chartexStyle: { lineHidden: true }, showMarker: false })],
      chartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(hidden.strokes).toContain('#A1B2C3');

    const permitted = recordingContext();
    renderChart(permitted.ctx, model({
      series: [series({ chartexStyle: { lineHidden: true }, showMarker: false })],
      chartStyleRoles: {
        dataPointLine: {
          lineColors: ['A1B2C3'],
          allowNoLineOverride: true,
        },
      },
    }), RECT, 1);
    expect(permitted.strokes).not.toContain('#A1B2C3');

    const numericOnly = recordingContext();
    renderChart(numericOnly.ctx, model({
      series: [series({ chartexStyle: { lineHidden: true }, showMarker: false })],
      classicChartStyleRoles: {
        dataPointLine: { lineColors: ['A1B2C3'] },
      },
      linkedChartStyleRoles: {},
    }), RECT, 1);
    expect(numericOnly.strokes).not.toContain('#A1B2C3');
  });

  it('does not let linked paint revive classic series or point lines removed by direct spPr', () => {
    const markerOnly = recordingContext();
    renderChart(markerOnly.ctx, model({
      chartType: 'scatter',
      scatterStyle: 'lineMarker',
      series: [series({
        categories: ['1', '2', '3'],
        lineHidden: true,
        showMarker: true,
      })],
      chartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
      linkedChartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(markerOnly.strokes).not.toContain('#A1B2C3');

    const borderlessSlice = recordingContext();
    renderChart(borderlessSlice.ctx, model({
      chartType: 'pie',
      varyColors: true,
      series: [series({
        values: [1, 2, 3],
        dataPointOverrides: [{ idx: 1, lineHidden: true }],
      })],
      chartStyleRoles: { dataPoint: { lineColors: ['A1B2C3'] } },
      linkedChartStyleRoles: { dataPoint: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    // The style still outlines the other two slices, but the direct no-line
    // point must remove one of the three linked outlines.
    expect(borderlessSlice.strokes.filter(color => color === '#A1B2C3')).toHaveLength(2);
  });

  it.each(['area', 'bar combo'] as const)(
    'uses dataPoint fill and outline for %s area geometry',
    family => {
      const rec = recordingContext();
      renderChart(rec.ctx, model(family === 'area' ? {
        chartType: 'area',
        chartStyleRoles: {
          dataPoint: { fillColors: ['135724'], lineColors: ['246813'] },
        },
      } : {
        chartType: 'clusteredBar',
        series: [
          series({ seriesType: 'area' }),
          series({ seriesType: 'bar', values: [2, 3, 4] }),
        ],
        chartStyleRoles: {
          dataPoint: { fillColors: ['135724'], lineColors: ['246813'] },
        },
      }), RECT, 1);

      expect(rec.fills).toContain('#135724');
      expect(rec.strokes).toContain('#246813');
    },
  );

  it.each(['bar', 'area'] as const)(
    'keeps a varyColors line overlay point-indexed in a %s combo',
    host => {
      const rec = recordingContext();
      const hostSeries = series({
        seriesType: host,
        values: [2, 3, 4],
      });
      const line = series({
        seriesType: 'line',
        values: [1, 2, 3],
        showMarker: false,
      });
      renderChart(rec.ctx, model({
        chartType: host === 'bar' ? 'clusteredBar' : 'area',
        series: [hostSeries, line],
        plotGroups: [
          {
            kind: host, seriesStart: 0, seriesCount: 1,
            categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
          },
          {
            kind: 'line', seriesStart: 1, seriesCount: 1, varyColors: true,
            categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
          },
        ],
        varyingPointChartStyleRolesByGroup: [
          null,
          {
            dataPointLine: {
              lineColors: ['AA0000', '00AA00', '0000AA'],
              lineFormattingIndices: [0, 1, 2],
            },
          },
        ],
      }), RECT, 1);

      expect(rec.strokes).toContain('#00AA00');
      expect(rec.strokes).toContain('#0000AA');
    },
  );

  it('uses dataPoint fill for filled radar polygons', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'radar',
      radarStyle: 'filled',
      chartStyleRoles: { dataPoint: { fillColors: ['ABC123'] } },
    }), RECT, 1);
    expect(rec.fills).toContain('#ABC123');
  });

  it('uses dataPoint and seriesLine roles for ofPie marks and connectors', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'ofPie',
      series: [series({ values: [40, 30, 20, 10] })],
      ofPie: {
        type: 'bar', splitType: 'pos', splitPos: 2,
        secondPieSizePercent: 75, gapWidthPercent: 100, seriesLines: true,
      },
      chartStyleRoles: {
        dataPoint: { fillColors: ['110000', '002200', '000033', '444400'] },
        seriesLine: { lineColors: ['778899'] },
      },
    }), RECT, 1);

    expect([...rec.fills, ...rec.rectFills]).toContain('#000033');
    expect(rec.strokes).toContain('#778899');
  });

  it('keeps a geometry-only point marker on its owning series style index', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'line',
      series: [series({
        chartexFormatIdx: 2,
        markerSymbol: 'circle',
        dataPointOverrides: [{ idx: 1, markerSize: 6 }],
      })],
      chartStyleRoles: {
        dataPointMarker: { fillColors: ['AA0000', '00AA00', '0000AA'] },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#0000AA')).toHaveLength(3);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(0);
  });

  it('separates bubble point and series style palette indexes', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'bubble',
      series: [series({
        chartexFormatIdx: 2,
        categories: ['1', '2'], values: [1, 2], bubbleSizes: [1, 1],
        chartexStyle: { fillColors: ['AA0000', '00AA00', '0000AA'] },
        dataPointOverrides: [{
          idx: 1,
          chartexStyle: { fillColors: ['CC0000', '00CC00', '0000CC'] },
        }],
      })],
    }), RECT, 1);

    expect(rec.fills).toEqual(['rgba(0,0,170,1)', 'rgba(0,204,0,1)']);
  });
});
