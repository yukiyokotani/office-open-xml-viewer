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
  markerStrokes: Array<{ color: string; width: number }>;
} {
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
  // Radial placement of `ctr` (and default) data labels, verified against the
  // sample-14.pdf ground truth (PowerPoint's own render):
  //   • SOLID pie   → labels sit near the rim at ≈0.88·outerR (measured 0.878 /
  //     0.888 / 0.887 / 0.912 across the 54/27/14/5% slices; center + outer
  //     radius from a least-squares rim fit, residual std 0.43pt). PowerPoint
  //     does NOT place a pie `ctr` label at the disc's geometric mid-radius.
  //   • DOUGHNUT    → labels sit at the RING midpoint (innerR+outerR)/2. For the
  //     55% hole this is 0.775·outerR, matching the PDF (measured 0.772–0.778
  //     across all five slices). This branch must not move.
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

  it('places SOLID-pie ctr labels near the rim (≈0.88R), not at mid-radius', () => {
    const rec = recordingCtx();
    renderChart(rec.ctx, pieModel('pie', 'ctr'), RECT, 1);
    const { cx, cy, outerR } = pieGeometry(rec.arcs);
    const labels = rec.texts.filter(t => /^\d+%$/.test(t.text));
    expect(labels.length).toBe(4);
    for (const l of labels) {
      const ratio = Math.hypot(l.x - cx, l.y - cy) / outerR;
      // Near the rim: comfortably beyond mid-radius, inside the outer edge.
      expect(ratio).toBeGreaterThan(0.8);
      expect(ratio).toBeLessThan(1.0);
      // And specifically close to the PowerPoint-measured ≈0.88 constant.
      expect(Math.abs(ratio - 0.88)).toBeLessThan(0.06);
    }
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
