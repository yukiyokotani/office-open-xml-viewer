import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import { renderChart } from './renderer.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';

function chart(over: Partial<ChartModel> = {}): ChartModel {
  return {
    chartType: 'clusteredBar',
    title: null,
    categories: ['A'],
    series: [{ name: 'S', color: null, values: [1] }],
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

function context() {
  const state: Record<string, unknown> = {
    canvas: { width: 0, height: 0 },
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    lineCap: 'butt',
    lineJoin: 'miter',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    globalAlpha: 1,
    globalCompositeOperation: 'source-over',
    shadowColor: 'transparent',
    shadowBlur: 0,
    shadowOffsetX: 0,
    shadowOffsetY: 0,
  };
  const saves: Array<Record<string, unknown>> = [];
  const fills: Array<{ shadow: string; x: number; y: number; w: number; h: number }> = [];
  const pathFills: string[] = [];
  const effectPathFills: Array<{ shadow: string; composite: string }> = [];
  const pathStrokes: Array<{ shadow: string; composite: string }> = [];
  let dash: number[] = [];
  const methods: Record<string, (...args: never[]) => unknown> = {
    save: () => saves.push({ ...state }),
    restore: () => Object.assign(state, saves.pop()),
    measureText: (text: string) => ({ width: text.length * 6 }),
    fillRect: (x: number, y: number, w: number, h: number) => fills.push({
      shadow: String(state.shadowColor), x, y, w, h,
    }),
    fill: () => {
      pathFills.push(String(state.shadowColor));
      effectPathFills.push({
        shadow: String(state.shadowColor),
        composite: String(state.globalCompositeOperation),
      });
    },
    stroke: () => pathStrokes.push({
      shadow: String(state.shadowColor),
      composite: String(state.globalCompositeOperation),
    }),
    getLineDash: () => [...dash],
    setLineDash: (value: number[]) => { dash = [...value]; },
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    createLinearGradient: () => ({ addColorStop() {} }),
    createRadialGradient: () => ({ addColorStop() {} }),
  } as unknown as Record<string, (...args: never[]) => unknown>;
  const noop = () => undefined;
  const ctx = new Proxy(state, {
    get(target, property: string) {
      if (property in methods) return methods[property];
      if (property in target) return target[property];
      return noop;
    },
    set(target, property: string, value) {
      target[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fills, pathFills, effectPathFills, pathStrokes };
}

describe('classic chart style effect renderer wiring', () => {
  const shadow = {
    color: '123456', alpha: 0.5, blur: 12_700, dist: 12_700, dir: 0,
  };

  it('paints the numeric dataPoint outer shadow on classic bars', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.fills.some(fill => fill.shadow === 'rgba(18,52,86,0.5)')).toBe(true);
  });

  it('keeps local and linked data-label effect palette domains distinct', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [{
        name: 'S', color: null, values: [null, 2],
        dataLabelOverrides: [{ idx: 1, text: '', showVal: true }],
      }],
      chartStyleRoles: {
        dataLabel: {
          fillColors: ['FFFFFF'], fillPaintAuthored: true,
          shadows: [
            { ...shadow, color: 'AA0000' },
            { ...shadow, color: '00BB00' },
          ],
          effectFormattingIndices: [0, 1],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.fills.some(fill => fill.shadow === 'rgba(0,187,0,0.5)')).toBe(true);
    expect(recorded.fills.some(fill => fill.shadow === 'rgba(170,0,0,0.5)')).toBe(false);
  });

  it('lets an authored empty direct effect list suppress the style shadow', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      series: [{
        name: 'S', color: null, values: [1],
        chartexStyle: { effectAuthored: true },
      }],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.fills.every(fill => fill.shadow === 'transparent')).toBe(true);
  });

  it('wires dataPointMarker effects to classic line markers', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      series: [{ name: 'S', color: null, values: [1], markerSymbol: 'circle' }],
      classicChartStyleRoles: {
        dataPointMarker: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathFills).toContain('rgba(18,52,86,0.5)');
  });

  it('routes automatic line markers through linked marker effects', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      series: [{ name: 'S', color: null, values: [1] }],
      classicChartStyleRoles: {
        dataPointMarker: {
          shadows: [shadow], effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathFills).toContain('rgba(18,52,86,0.5)');
  });

  it('paints data-point effects on 2-D legend swatch bodies', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      showLegend: true,
      legendPos: 'r',
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['4472C4'], fillPaintAuthored: true,
          shadows: [shadow], effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.fills.filter(fill =>
      fill.shadow === 'rgba(18,52,86,0.5)'
    ).length).toBeGreaterThanOrEqual(2);
  });

  it('resolves linked marker effects for a series-driven legend key', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line', showLegend: true, legendPos: 'r',
      series: [{ name: 'S', color: null, values: [] }],
      classicChartStyleRoles: {
        dataPointMarker: { shadows: [shadow], effectAuthored: true },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathFills).toContain('rgba(18,52,86,0.5)');
  });

  it('wires dataPointLine effects to a non-varying bar-combo line', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'clusteredBar',
      categories: ['A', 'B'],
      series: [
        { name: 'Bars', color: null, values: [2, 3] },
        {
          name: 'Line', color: null, values: [1, 4],
          seriesType: 'line', showMarker: false,
        },
      ],
      classicChartStyleRoles: {
        dataPointLine: {
          lineColors: ['4472C4'],
          linePaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathStrokes.some(stroke =>
      stroke.shadow === 'rgba(18,52,86,0.5)'
    )).toBe(true);
  });

  it('lets direct dPt effects style a line point when marker spPr omits effects', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      series: [{
        name: 'S', color: null, values: [1], markerSymbol: 'circle',
        dataPointOverrides: [{
          idx: 0,
          chartexStyle: { shadows: [shadow], effectAuthored: true },
        }],
      }],
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathFills).toContain('rgba(18,52,86,0.5)');
  });

  it('wires dataPoint effects to classic pie slices', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'pie',
      series: [{ name: 'S', color: null, values: [1] }],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.pathFills).toContain('rgba(18,52,86,0.5)');
  });

  it.each([false, true])(
    'selects varying bubble effects by point index (bubble3D=%s)',
    bubble3D => {
      const recorded = context();
      const pointShadows = ['AA0000', '00BB00'].map(color => ({
        ...shadow,
        color,
      }));
      const role = {
        fillColors: ['4472C4', 'ED7D31'],
        fillPaintAuthored: true,
        shadows: pointShadows,
        effectFormattingIndices: [0, 1],
        effectAuthored: true,
      };
      renderChart(recorded.ctx, chart({
        chartType: 'bubble',
        categories: ['1', '2'],
        series: [{
          name: 'S', color: null, values: [1, 2], bubbleSizes: [100, 100],
          bubble3D,
        }],
        varyingPointChartStyleRoles: {
          dataPoint: role,
          dataPoint3D: role,
        },
      }), { x: 0, y: 0, w: 320, h: 180 }, 1);
      expect(recorded.pathFills).toContain('rgba(170,0,0,0.5)');
      expect(recorded.pathFills).toContain('rgba(0,187,0,0.5)');
    },
  );

  it('wires upBar and downBar effects to line-group decorations', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [
        { name: 'Open', color: null, values: [1, 3], lineGroupIndex: 0, showMarker: false },
        { name: 'Close', color: null, values: [2, 2], lineGroupIndex: 0, showMarker: false },
      ],
      lineGroupDecorations: [{
        groupIndex: 0,
        upDownBars: { gapWidthPercent: 150, up: {}, down: {} },
      }],
      classicChartStyleRoles: {
        upBar: {
          fillColors: ['FFFFFF'], fillPaintAuthored: true,
          shadows: [shadow], effectAuthored: true,
        },
        downBar: {
          fillColors: ['000000'], fillPaintAuthored: true,
          shadows: [shadow], effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1);
    expect(recorded.fills.filter(fill => fill.shadow === 'rgba(18,52,86,0.5)')).toHaveLength(2);
  });

  it('composites dataPoint3D outer shadows behind the depth-sorted scene', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      classicChartStyleRoles: {
        dataPoint3D: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1, 0, {
      render: renderSimpleThreeDChart,
    });
    expect(recorded.effectPathFills.some(fill =>
      fill.shadow === 'rgba(18,52,86,0.5)'
        && fill.composite === 'destination-over')).toBe(true);
  });

  it('wires dataPointMarker effects to projected 3-D line markers', () => {
    const recorded = context();
    renderChart(recorded.ctx, chart({
      chartType: 'line',
      categories: ['A', 'B'],
      series: [{
        name: 'S', color: null, values: [1, 2],
        markerSymbol: 'circle', showMarker: true,
      }],
      threeD: { rotationX: 15, rotationY: 20, depthPercent: 100, perspective: 30 },
      classicChartStyleRoles: {
        dataPointMarker: {
          fillColors: ['4472C4'],
          fillPaintAuthored: true,
          shadows: [shadow],
          effectAuthored: true,
        },
      },
    }), { x: 0, y: 0, w: 320, h: 180 }, 1, 0, {
      render: renderSimpleThreeDChart,
    });
    expect(recorded.effectPathFills.some(fill =>
      fill.shadow === 'rgba(18,52,86,0.5)'
        && fill.composite === 'source-over')).toBe(true);
  });

  it.each(['chartArea', 'plotArea', 'legend'] as const)(
    'wires %s frame paint and effects through the shared style role',
    role => {
      const recorded = context();
      renderChart(recorded.ctx, chart({
        showLegend: role === 'legend',
        legendPos: role === 'legend' ? 'r' : null,
        classicChartStyleRoles: {
          [role]: {
            fillColors: ['DDEEFF'],
            fillPaintAuthored: true,
            shadows: [shadow],
            effectAuthored: true,
          },
        },
      }), { x: 0, y: 0, w: 320, h: 180 }, 1);
      expect(recorded.fills.some(fill =>
        fill.shadow === 'rgba(18,52,86,0.5)'
      )).toBe(true);
    },
  );

  it.each([
    ['line', chart({
      chartType: 'line', categories: ['A', 'B'],
      series: [{ name: 'S', color: null, values: [1, 2], showMarker: false }],
    }), 'dataPointLine'],
    ['area', chart({
      chartType: 'area', categories: ['A', 'B', 'C'],
      series: [{ name: 'S', color: null, values: [1, 2, 1] }],
    }), 'dataPoint'],
    ['radar', chart({
      chartType: 'radar', radarStyle: 'standard', categories: ['A', 'B', 'C'],
      series: [{ name: 'S', color: null, values: [1, 2, 1], showMarker: false }],
    }), 'dataPointLine'],
    ['scatter', chart({
      chartType: 'scatter', scatterStyle: 'line', categories: ['0', '1'],
      series: [{ name: 'S', color: null, categories: ['0', '1'], values: [1, 2] }],
    }), 'dataPointLine'],
    ['stock', chart({
      chartType: 'stock', categories: ['A', 'B'],
      series: [
        { name: 'High', color: null, values: [3, 4] },
        { name: 'Low', color: null, values: [1, 2] },
        { name: 'Close', color: null, values: [2, 3] },
      ],
    }), 'dataPointLine'],
  ] as const)('wires numeric effects to %s series bodies', (_name, model, role) => {
    const recorded = context();
    model.classicChartStyleRoles = {
      [role]: {
        shadows: [shadow],
        effectAuthored: true,
      },
    };
    renderChart(recorded.ctx, model, { x: 0, y: 0, w: 320, h: 180 }, 1);
    const effects = [...recorded.effectPathFills, ...recorded.pathStrokes];
    expect(effects.some(paint => paint.shadow === 'rgba(18,52,86,0.5)')).toBe(true);
  });
});
