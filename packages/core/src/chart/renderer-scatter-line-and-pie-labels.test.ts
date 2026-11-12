// Regression tests for two chart data-fidelity fixes:
//
//   1. Scatter connecting line: a series-level `<c:spPr><a:ln><a:noFill/>`
//      (ECMA-376 §21.2.2.198) OVERRIDES the chart-group `<c:scatterStyle>`
//      (§21.2.2.42). Excel draws a MARKERS-ONLY scatter when the series line is
//      `<a:noFill/>` even though the group default is `lineMarker`. The parser
//      records this as `ChartSeries.lineHidden`; the renderer must draw no
//      connecting line for such a series (sample-30 sheet 1's two scatters).
//
//   2. Pie per-point data labels: a per-point `<c:dLbl idx>` (§21.2.2.47)
//      carries its own show-flag group (§21.2.2.177/.180/.187/.189) and text
//      style, which OVERRIDE the series-level `<c:dLbls>` defaults (§21.2.2.49)
//      for that one slice. sample-14 slide-7's pie sets `showCatName=0
//      showPercent=1` + white text PER SLICE while the series default is
//      `showCatName=1` black — so each label must render as WHITE PERCENT-ONLY,
//      not black "category + percent".
//
// Both broke when the pptx/xlsx parsers were unified onto the shared
// ooxml-common `parse_chart_part` (which now surfaces scatterStyle /
// series-level dLbls the renderer honors) without carrying the override /
// noFill semantics that suppress or reshape those defaults.

import { describe, it, expect } from 'vitest';
import type { ChartModel, ChartSeries, ChartRect } from '../types/chart';
import { renderChart } from './renderer.js';

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };

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

interface TextCall {
  text: string;
  fillStyle: string;
  x: number;
  y: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
}

/** Recording context that captures: (a) `stroke()` calls that follow a
 *  `moveTo`/`lineTo` path with ≥2 vertices AND whose strokeStyle is `matchColor`
 *  (the series line color) — this isolates the scatter connecting line from
 *  gridlines/axis rules, which stroke in grey; (b) `fillText` calls with the
 *  fillStyle in effect. Enough to assert the scatter line pass and pie labels. */
interface ArcCall { cx: number; cy: number; r: number }

function recordingCtx(matchColor = '#4f81bd'): {
  ctx: CanvasRenderingContext2D;
  counts: { polylineStrokes: number };
  texts: TextCall[];
  arcs: ArcCall[];
  rects: Array<{ kind: string; color: string; x: number; y: number; w: number; h: number }>;
  markerStrokes: Array<{ color: string; width: number }>;
} {
  const rects: Array<{ kind: string; color: string; x: number; y: number; w: number; h: number }> = [];
  const texts: TextCall[] = [];
  const arcs: ArcCall[] = [];
  const markerStrokes: Array<{ color: string; width: number }> = [];
  const counts = { polylineStrokes: 0 };
  let pathVerts = 0;
  let pathHasArc = false;
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
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_t, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'fillRect':
        case 'strokeRect':
          return (x: number, y: number, w: number, h: number) => rects.push({ kind: prop, color: String(state[prop === 'fillRect' ? 'fillStyle' : 'strokeStyle']), x, y, w, h });
        case 'measureText':
          return (t: string) => {
            const px = fontPx(String(state.font));
            let w = 0;
            for (const ch of String(t)) w += ch.charCodeAt(0) > 0x2e7f ? px : px * 0.6;
            return { width: w };
          };
        case 'beginPath':
          return () => { pathVerts = 0; pathHasArc = false; };
        // Pie/doughnut slices are drawn with `arc(cx, cy, r, …)`; recording the
        // centre + radius lets a test recover the exact pie geometry the renderer
        // used (the outermost rim is the max-radius arc) without duplicating the
        // frame math.
        case 'arc':
          return (cx: number, cy: number, r: number) => {
            pathHasArc = true;
            arcs.push({ cx, cy, r });
          };
        case 'moveTo':
        case 'lineTo':
        case 'bezierCurveTo':
        case 'quadraticCurveTo':
          return () => { pathVerts += 1; };
        // A marker is a single `arc` + `fill`; it does not build a multi-vertex
        // moveTo/lineTo path. Only a stroke in the SERIES color with ≥2 vertices
        // is a connecting line (gridlines/axes stroke grey → excluded).
        case 'stroke':
          return () => {
            if (pathHasArc) {
              markerStrokes.push({ color: String(state.strokeStyle), width: Number(state.lineWidth) });
            }
            if (pathVerts >= 2 && String(state.strokeStyle).toLowerCase() === matchColor.toLowerCase()) {
              counts.polylineStrokes += 1;
            }
          };
        case 'fillText':
          return (text: string, x: number, y: number) =>
            texts.push({
              text,
              fillStyle: String(state.fillStyle),
              x,
              y,
              font: String(state.font),
              textAlign: state.textAlign as CanvasTextAlign,
              textBaseline: state.textBaseline as CanvasTextBaseline,
            });
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
    counts,
    rects,
    texts,
    arcs,
    markerStrokes,
  };
}

/** Recover the pie/doughnut centre + outer radius from recorded `arc()` calls.
 *  The outermost rim is the arc with the largest radius; every slice on that ring
 *  shares the same centre. */
function pieGeometry(arcs: ArcCall[]): { cx: number; cy: number; outerR: number } {
  const outer = arcs.reduce((a, b) => (b.r > a.r ? b : a));
  return { cx: outer.cx, cy: outer.cy, outerR: outer.r };
}

describe('scatter series-line noFill overrides scatterStyle (§21.2.2.198)', () => {
  // sample-30 sheet 1: `<c:scatterStyle val="lineMarker">` at the group level,
  // but each series carries `<c:spPr><a:ln><a:noFill/></c:spPr>` → markers only.
  const scatterModel = (lineHidden: boolean | null): ChartModel =>
    baseModel({
      chartType: 'scatter',
      scatterStyle: 'lineMarker',
      categories: ['1', '2', '3', '4'],
      series: [series({
        name: 'S',
        color: '4f81bd',
        values: [10, 20, 15, 25],
        categories: ['1', '2', '3', '4'],
        showMarker: true,
        lineHidden,
      })],
    });

  it('draws NO connecting polyline when the series line is noFill', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, scatterModel(true), RECT, 1);
    expect(rec.counts.polylineStrokes).toBe(0);
  });

  it('still draws the connecting polyline for a normal lineMarker scatter', () => {
    // Guard: the fix must not suppress lines universally — a series WITHOUT the
    // noFill flag keeps the group scatterStyle line (regression the other way).
    const rec = recordingCtx();
    renderChart(rec.ctx, scatterModel(null), RECT, 1);
    expect(rec.counts.polylineStrokes).toBeGreaterThan(0);
  });
});

describe('bubble series shape outline (§21.2.2.198)', () => {
  it('uses the series spPr line color and authored EMU width', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['1'],
      series: [series({
        color: '4F81BD66',
        lineColor: '4F81BD',
        lineWidthEmu: 9525,
        categories: ['1'],
        values: [2],
        bubbleSizes: [100],
      })],
    }), RECT, 4 / 3);

    expect(rec.markerStrokes).toContainEqual({ color: '#4F81BD', width: 1 });
  });
});

describe('pie per-point dLbl flags override series defaults (§21.2.2.47)', () => {
  // sample-14 slide-7 pie: series default `<c:dLbls>` says showCatName=1,
  // black; every slice's `<c:dLbl>` says showCatName=0 showPercent=1, WHITE.
  // Result must be white percent-only labels.
  const pieModel = (): ChartModel =>
    baseModel({
      chartType: 'pie',
      showDataLabels: true,
      categories: ['Alpha', 'Beta', 'Gamma'],
      series: [series({
        name: 'Share',
        values: [50, 30, 20],
        categories: ['Alpha', 'Beta', 'Gamma'],
        dataPointColors: ['ff0000', '00ff00', '0000ff'],
        seriesDataLabels: {
          showVal: false,
          showCatName: true,   // series default WOULD show category names…
          showSerName: false,
          showPercent: true,
          fontColor: '000000', // …in black.
          fontBold: true,
          fontSizeHpt: 1800,
          formatCode: '0%',
          position: 'ctr',
        },
        dataLabelOverrides: [0, 1, 2].map(idx => ({
          idx,
          text: '',            // no custom text → compose from flags
          fontColor: 'FFFFFF', // per-slice WHITE
          fontBold: true,
          fontSizeHpt: 1200,
          showVal: false,
          showCatName: false,  // per-slice: percent only
          showSerName: false,
          showPercent: true,
        })),
      })],
    });

  it('renders white, percent-only labels (no category names, not black)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel(), RECT, 1);
    const labels = rec.texts.filter(t => /%$/.test(t.text) || /Alpha|Beta|Gamma/.test(t.text));
    expect(labels.length).toBeGreaterThan(0);
    // Every drawn slice label is a bare percentage…
    for (const l of labels) {
      expect(l.text).toMatch(/^\d+%$/);
      // …in white (the per-point override color), never the series-default black.
      expect(l.fillStyle.toLowerCase()).toBe('#ffffff');
    }
    // And no category name leaked through.
    expect(rec.texts.some(t => /Alpha|Beta|Gamma/.test(t.text))).toBe(false);
  });

  it('a genuine <c:delete> per-point override still skips that slice', () => {
    const model = pieModel();
    model.series[0].dataLabelOverrides = [
      { idx: 0, text: '', deleted: true },
      ...([1, 2].map(idx => ({
        idx, text: '', fontColor: 'FFFFFF', showCatName: false, showPercent: true,
      }))),
    ];
    const rec = recordingCtx();
    renderChart(rec.ctx, model, RECT, 1);
    const pctLabels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    // Two slices labeled (idx 1,2); the deleted idx 0 is skipped.
    expect(pctLabels.length).toBe(2);
  });
});

describe('pie / doughnut ctr data-label radius (§21.2.2.48, PowerPoint layout)', () => {
  // `ctr` centers labels halfway along the radial data marker. Unlike inEnd,
  // it does not require the text rectangle to fit inside the wedge: even a
  // narrow slice keeps its label. Check the production painter, not a helper.
  const pieModel = (chartType: 'pie' | 'doughnut', position?: string): ChartModel =>
    baseModel({
      chartType,
      showDataLabels: true,
      holeSize: chartType === 'doughnut' ? 55 : undefined,
      categories: ['A', 'B', 'C', 'D'],
      series: [series({
        name: 'Share',
        values: [54, 27, 14, 5],
        categories: ['A', 'B', 'C', 'D'],
        dataPointColors: ['ff0000', '00ff00', '0000ff', 'ffff00'],
        seriesDataLabels: {
          showVal: false, showCatName: false, showSerName: false, showPercent: true,
          fontColor: 'FFFFFF', fontBold: true, formatCode: '0%',
          ...(position ? { position } : {}),
        },
      })],
    });

  it('centers solid-pie labels at half-radius, including narrow slices', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel('pie', 'ctr'), RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const labels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    expect(labels.length).toBe(4);
    for (const l of labels) {
      const ratio = Math.hypot(l.x - cx, l.y - cy) / outerR;
      expect(ratio).toBeCloseTo(0.5, 5);
    }
  });

  it.each([false, true])('keeps ctr labels visible across a 5-degree wedge (boxed=%s)', boxed => {
    const model = pieModel('pie', 'inEnd');
    const s = model.series[0];
    s.values = [5, 355];
    s.dataPointOverrides = [{ idx: 0, explosion: 25 }];
    s.dataLabelOverrides = [{ idx: 0, text: '', position: 'ctr' }];
    s.seriesDataLabels = {
      ...s.seriesDataLabels,
      showVal: true, showPercent: false, showCatName: false, showSerName: false,
      formatCode: '0', fontSizeHpt: 1200,
      ...(boxed ? { labelBox: { fill: 'FFFFFF' } } : {}),
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, model, RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const label = rec.texts.find(t => t.text === '5');
    expect(label).toBeDefined();
    expect(Math.hypot((label as TextCall).x - cx, (label as TextCall).y - cy) / outerR)
      .toBeCloseTo(0.5, 5);
    // The point override must not change its sibling's inside-end placement.
    const sibling = rec.texts.find(t => t.text === '355') as TextCall;
    expect(sibling).toBeDefined();
    expect(Math.hypot(sibling.x - cx, sibling.y - cy) / outerR).toBeGreaterThan(0.8);
  });

  it('does not create a centered label for a zero-area slice', () => {
    const model = pieModel('pie', 'ctr');
    model.series[0].values = [0, 360];
    model.series[0].seriesDataLabels = {
      showVal: true, showPercent: false, showCatName: false, showSerName: false,
      position: 'ctr', formatCode: '0', fontSizeHpt: 1200,
    };
    const rec = recordingCtx();
    renderChart(rec.ctx, model, RECT, 1);
    expect(rec.texts.map(t => t.text)).toEqual(['360']);
  });

  it('keeps DOUGHNUT ctr labels at the ring midpoint (0.775R for a 55% hole)', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel('doughnut'), RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const labels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    expect(labels.length).toBeGreaterThan(0);
    for (const l of labels) {
      const ratio = Math.hypot(l.x - cx, l.y - cy) / outerR;
      // (innerR + outerR)/2 with innerR = 0.55·outerR ⇒ 0.775.
      expect(Math.abs(ratio - 0.775)).toBeLessThan(0.02);
    }
  });

  it('still pushes outEnd pie labels OUTSIDE the rim', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel('pie', 'outEnd'), RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const labels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    expect(labels.length).toBe(4);
    for (const l of labels) {
      const ratio = Math.hypot(l.x - cx, l.y - cy) / outerR;
      expect(ratio).toBeGreaterThan(1.0);
    }
  });

  it('keeps automatic outEnd label anchors beyond the pie rim', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Long right-side category', 'Long left-side category'],
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.3, y: 0.15, w: 0.4, h: 0.7,
      },
      series: [series({
        name: 'Share',
        values: [50, 50],
        categories: ['Long right-side category', 'Long left-side category'],
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd', separator: '\n', fontSizeHpt: 1200,
        },
      })],
    }), RECT, 1);

    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const categoryLabels = rec.texts.filter(t => t.text.startsWith('Long '));
    expect(categoryLabels).toHaveLength(2);
    for (const label of categoryLabels) {
      expect(Math.hypot(label.x - cx, label.y - cy)).toBeGreaterThan(outerR);
    }
  });

  it('keeps the complete outEnd text boxes outside the pie, not only their centers', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Long right-side category', 'Long left-side category'],
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.3, y: 0.15, w: 0.4, h: 0.7,
      },
      series: [series({
        name: 'Share',
        values: [50, 50],
        categories: ['Long right-side category', 'Long left-side category'],
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd', separator: '\n', fontSizeHpt: 1200,
        },
      })],
    }), RECT, 1);

    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const categoryLabels = rec.texts.filter(t => t.text.startsWith('Long '));
    expect(categoryLabels).toHaveLength(2);
    for (const label of categoryLabels) {
      const fontPx = Number(/(\d+(?:\.\d+)?)px/.exec(label.font)?.[1] ?? 12);
      const halfW = label.text.length * fontPx * 0.6 / 2;
      const halfH = fontPx / 2;
      const dx = Math.max(Math.abs(label.x - cx) - halfW, 0);
      const dy = Math.max(Math.abs(label.y - cy) - halfH, 0);
      expect(Math.hypot(dx, dy)).toBeGreaterThanOrEqual(outerR);
    }
  });

  it('does not invent leader lines when ordinary outEnd labels do not collide', () => {
    const rec = recordingCtx('#A6A6A6');
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Fares and taxes', 'Local funds', 'State funds', 'Federal funds'],
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.288247, y: 0.186997, w: 0.402462, h: 0.670347,
      },
      series: [series({
        name: 'Share',
        values: [43, 27, 23, 7],
        categories: ['Fares and taxes', 'Local funds', 'State funds', 'Federal funds'],
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd', separator: '\n', fontSizeHpt: 1100,
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
        },
      })],
    }), { x: 0, y: 0, w: 740, h: 545 }, 4 / 3);

    expect(rec.counts.polylineStrokes).toBe(0);
  });

  it('does not treat a transparent data-label spPr as a visible callout box', () => {
    const rec = recordingCtx('#A6A6A6');
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['Fares and taxes', 'Local funds', 'State funds', 'Federal funds'],
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.288247, y: 0.186997, w: 0.402462, h: 0.670347,
      },
      series: [series({
        name: 'Share',
        values: [43, 27, 23, 7],
        categories: ['Fares and taxes', 'Local funds', 'State funds', 'Federal funds'],
        seriesDataLabels: {
          showVal: true, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd', separator: '\n', fontSizeHpt: 1100,
          showLeaderLines: true, leaderLineColor: 'A6A6A6', leaderLineWidthEmu: 9525,
          labelBox: {
            fillHidden: true,
            fillPaintAuthored: true,
            borderHidden: true,
            borderPaintAuthored: true,
          },
        },
      })],
    }), { x: 0, y: 0, w: 740, h: 545 }, 4 / 3);

    expect(rec.counts.polylineStrokes).toBe(0);
  });

  it('does not synthesize leader lines by moving crowded automatic outEnd labels', () => {
    const rec = recordingCtx('#777777');
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['First narrow slice', 'Second narrow slice', 'Remainder'],
      series: [series({
        name: 'Share',
        values: [1, 1, 98],
        categories: ['First narrow slice', 'Second narrow slice', 'Remainder'],
        seriesDataLabels: {
          showVal: false, showCatName: true, showSerName: false, showPercent: false,
          position: 'outEnd', fontSizeHpt: 1200,
          showLeaderLines: true, leaderLineColor: '777777', leaderLineWidthEmu: 12700,
        },
      })],
    }), RECT, 1);

    expect(rec.counts.polylineStrokes).toBe(0);
  });
});

describe('authored radial and bubble geometry', () => {
  it('honors plotArea manualLayout for pie center and radius', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie',
      categories: ['A', 'B'],
      series: [series({ values: [60, 40] })],
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.2, w: 0.4, h: 0.6,
      },
    }), RECT, 1);

    const geometry = pieGeometry(rec.arcs);
    expect(geometry.cx).toBeCloseTo(RECT.w * 0.3);
    expect(geometry.cy).toBeCloseTo(RECT.h * 0.5);
    expect(geometry.outerR).toBeCloseTo(Math.min(RECT.w * 0.4, RECT.h * 0.6) * 0.42);
  });

  it('applies Office\'s bounded bubbleScale curve to the shorter plot dimension', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'bubble',
      categories: ['0'],
      series: [series({ values: [0], bubbleSizes: [100] })],
      bubbleScale: 40,
      catAxisMin: -1,
      catAxisMax: 1,
      valMin: -1,
      valMax: 1,
      plotAreaManualLayout: {
        xMode: 'edge', yMode: 'edge', x: 0.25, y: 0.25, w: 0.5, h: 0.5,
      },
    }), RECT, 1);

    const marker = rec.arcs.at(-1);
    // Plot is 320×180. Excel's curve is shortSide*scale/(300+scale), so an
    // authored bubbleScale=40 gives a 21.176px diameter / 10.588px radius.
    expect(marker?.r).toBeCloseTo(10.588235, 4);
  });
});


describe('PowerPoint radial plot frame', () => {
  it.each(['pie', 'doughnut'])('fits the automatic %s frame to the circle and preserves authored paint', chartType => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType, powerpointRadialPlotFrame: true,
      plotAreaBg: '123456', plotAreaLineColor: 'ABCDEF', plotAreaLineWidthEmu: 12700,
      series: [series({ values: [3, 7, 5] })],
    }), RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const fill = rec.rects.find(r => r.color === '#123456');
    expect(fill).toMatchObject({ x: cx - outerR, y: cy - outerR, w: outerR * 2, h: outerR * 2 });
    const line = rec.rects.find(r => r.color === '#ABCDEF');
    expect(line).toMatchObject({ x: cx - outerR + 0.5, y: cy - outerR + 0.5, w: outerR * 2 - 1, h: outerR * 2 - 1 });
  });

  it('retains the automatic rectangle for hosts without the PowerPoint policy', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'pie', plotAreaBg: '123456', series: [series({ values: [3, 7, 5] })],
    }), RECT, 1);
    const fill = rec.rects.find(r => r.color === '#123456');
    expect(fill?.w).toBe(RECT.w);
    expect(fill?.h).toBeGreaterThan(pieGeometry(rec.arcs).outerR * 2);
  });

  it.each([false, true])('preserves the manual frame with PowerPoint policy=%s', powerpointRadialPlotFrame => {
    const rec = recordingCtx();
    renderChart(rec.ctx, baseModel({
      chartType: 'doughnut', powerpointRadialPlotFrame,
      plotAreaBg: '123456',
      plotAreaManualLayout: { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.2, w: 0.7, h: 0.6 },
      series: [series({ values: [3, 7, 5] })],
    }), RECT, 1);
    expect(rec.rects.find(r => r.color === '#123456')).toMatchObject({ x: 64, y: 72, w: 448, h: 216 });
  });
});
