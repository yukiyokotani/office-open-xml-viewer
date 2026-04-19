// Unified chart renderer. Dispatches on canonical `ChartModel.chartType` and
// delegates to per-family implementations (bar, line, area, pie, radar,
// scatter, waterfall). Ported from the xlsx implementation with pptx
// extensions (valMin-aware axis, plotAreaBg, dataPointColors, waterfall).

import type { ChartModel, ChartRect, ChartSeries } from '../types/chart';

// ─── Palette + helpers ──────────────────────────────────────────────────────

const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

function chartColor(idx: number, series?: { color?: string | null } | null): string {
  if (series?.color) return `#${series.color}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

function pieSliceColor(idx: number, series: ChartSeries): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

function hexToRgba(hex: string, alpha: number): string {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

function niceAxisMax(dataMax: number, step: number): number {
  if (dataMax <= 0) return step;
  const ax = Math.ceil(dataMax / step) * step;
  return Math.abs(ax - dataMax) < step * 1e-9 ? ax + step : ax;
}

function niceAxisMin(dataMin: number, step: number): number {
  if (dataMin >= 0) return 0;
  const ax = Math.floor(dataMin / step) * step;
  return Math.abs(ax - dataMin) < step * 1e-9 ? ax - step : ax;
}

function formatChartVal(v: number): string {
  if (Math.abs(v) >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`;
  if (Math.abs(v) >= 1_000) return `${(v / 1_000).toFixed(1)}k`;
  return Number.isInteger(v) ? String(v) : v.toFixed(2).replace(/\.?0+$/, '');
}

function drawAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  px0: number, py0: number, pw: number, ph: number,
  axis: 'cat' | 'val',
  fontSize: number,
): void {
  ctx.save();
  ctx.font = `${fontSize}px sans-serif`;
  ctx.fillStyle = '#555';
  if (axis === 'cat') {
    ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
    ctx.fillText(text.slice(0, 30), px0 + pw / 2, py0 + ph + fontSize + 2);
  } else {
    ctx.translate(px0 - fontSize - 4, py0 + ph / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(text.slice(0, 30), 0, 0);
  }
  ctx.restore();
}

function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartSeries[],
  lx: number, ly: number, lw: number, lh: number,
): void {
  const fontSize = Math.max(9, Math.min(12, lh / (series.length + 1)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const sw = 10; const gap = 4;
  const rowH = fontSize + 4;
  let ry = ly + (lh - rowH * series.length) / 2;
  for (let i = 0; i < series.length; i++) {
    ctx.fillStyle = chartColor(i, series[i]);
    ctx.fillRect(lx, ry, sw, fontSize);
    ctx.fillStyle = '#333'; ctx.textAlign = 'left';
    const label = series[i].name || `Series ${i + 1}`;
    ctx.fillText(label.slice(0, 20), lx + sw + gap, ry + fontSize / 2);
    ry += rowH;
  }
  void lw;
}

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  title: string | null,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!title) return;
  ctx.font = `bold ${fontSize}px sans-serif`;
  ctx.fillStyle = '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(title, x + w / 2, y);
}

// ─── Category helper ────────────────────────────────────────────────────────

function chartCategories(chart: ChartModel): string[] {
  if (chart.categories.length > 0) return chart.categories;
  const first = chart.series[0];
  return first?.categories ?? [];
}

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

function renderBarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const barSeries  = chart.series.filter(s => s.seriesType !== 'line');
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;

  const titleH   = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW  = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = {
    t: titleH + h * 0.04,
    r: legendW + w * 0.03,
    b: h * 0.14 + catTitleH,
    l: w * 0.12 + valTitleW,
  };
  if (isH) { pad.l = w * 0.22 + valTitleW; pad.b = h * 0.08 + catTitleH; }

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw  = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    let stackSum = 0;
    for (const s of barSeries) {
      const v = s.values[ci] ?? 0;
      if (stacked) stackSum += Math.abs(v);
      else dataMax = Math.max(dataMax, Math.abs(v));
    }
    if (stacked) dataMax = Math.max(dataMax, stackSum);
  }
  if (pct) dataMax = 100;
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;

  const step  = niceStep(dataMax);
  const axMax = chart.valMax ?? niceAxisMax(dataMax, step);

  const gridColor = '#e0e0e0';
  const steps = Math.round(axMax / step);
  ctx.textBaseline = 'middle';
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  ctx.fillStyle = '#555';

  if (!chart.valAxisHidden) {
    for (let si = 0; si <= steps; si++) {
      const val = si * step;
      const label = pct ? `${Math.round(val)}%` : val >= 1000 ? `${(val / 1000).toFixed(1)}k` : String(val);
      if (!isH) {
        const gy = py0 + ph - (val / axMax) * ph;
        ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
        ctx.lineWidth = si === 0 ? 1 : 0.5;
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        ctx.textAlign = 'right';
        ctx.fillText(label, px0 - 4, gy);
      } else {
        const gx = px0 + (val / axMax) * pw;
        ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
        ctx.lineWidth = si === 0 ? 1 : 0.5;
        ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
        ctx.textAlign = 'center';
        ctx.fillText(label, gx, py0 + ph + 10);
      }
    }
  }

  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  if (!isH) {
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  } else {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  const catGap = !isH ? pw / n : ph / n;
  const barW   = catGap * (stacked ? 0.6 : 0.6 / Math.max(1, barSeries.length));
  const clusterGap = stacked ? 0 : catGap * 0.6 / Math.max(1, barSeries.length);
  const catStart   = catGap * 0.2;

  for (let ci = 0; ci < n; ci++) {
    let stackOffset = 0;
    let stackSum = 0;
    if (pct) {
      for (const s of barSeries) stackSum += Math.abs(s.values[ci] ?? 0);
      if (stackSum === 0) stackSum = 1;
    }

    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const raw = s.values[ci] ?? 0;
      const val = pct ? (Math.abs(raw) / stackSum) * 100 : Math.abs(raw);
      const color = chartColor(si, s);

      if (!isH) {
        const bx = stacked
          ? px0 + ci * catGap + catStart
          : px0 + ci * catGap + catStart + si * clusterGap;
        const barH = (val / axMax) * ph;
        const by   = py0 + ph - (stacked ? (stackOffset + val) : val) / axMax * ph;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barW, barH);
        if (chart.showDataLabels && val > 0) {
          const lsz = Math.max(7, Math.min(10, barW * 0.7));
          ctx.font = `bold ${lsz}px sans-serif`;
          ctx.fillStyle = '#333'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
          ctx.fillText(pct ? `${Math.round(val)}%` : formatChartVal(val), bx + barW / 2, by - 1);
        }
      } else {
        const by = stacked
          ? py0 + (n - 1 - ci) * catGap + catStart
          : py0 + (n - 1 - ci) * catGap + catStart + si * clusterGap;
        const barL = (val / axMax) * pw;
        const bx   = stacked ? px0 + (stackOffset / axMax) * pw : px0;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barL, barW);
        if (chart.showDataLabels && val > 0) {
          const lsz = Math.max(7, Math.min(10, barW * 0.7));
          ctx.font = `bold ${lsz}px sans-serif`;
          ctx.fillStyle = '#333'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
          ctx.fillText(pct ? `${Math.round(val)}%` : formatChartVal(val), bx + barL + 2, by + barW / 2);
        }
      }
      if (stacked) stackOffset += val;
    }
  }

  if (!chart.catAxisHidden) {
    ctx.fillStyle = '#555';
    ctx.font = `${Math.max(8, Math.min(11, catGap * 0.5))}px sans-serif`;
    for (let ci = 0; ci < n; ci++) {
      const label = (cats[ci] ?? '').toString().slice(0, 12);
      if (!isH) {
        const lx = px0 + ci * catGap + catGap / 2;
        ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(label, lx, py0 + ph + 3);
      } else {
        const ly = py0 + (n - 1 - ci) * catGap + catGap / 2;
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(label, px0 - 4, ly);
      }
    }
  }

  if (lineSeries.length > 0 && !isH) {
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const color = chartColor(barSeries.length + si, s);
      ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.setLineDash([]);
      ctx.beginPath();
      let started = false;
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) { started = false; continue; }
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = py0 + ph - (v / axMax) * ph;
        if (!started) { ctx.moveTo(lx, ly); started = true; } else ctx.lineTo(lx, ly);
      }
      ctx.stroke();
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) continue;
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = py0 + ph - (v / axMax) * ph;
        ctx.fillStyle = color;
        ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
      }
    }
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function renderLineChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;

  const titleH   = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW  = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = { t: titleH + h * 0.04, r: legendW + w * 0.05, b: h * 0.14 + catTitleH, l: w * 0.12 + valTitleW };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  let dataMin = Infinity; let dataMax = -Infinity;
  for (const s of chart.series) for (const v of s.values) if (v != null) { dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v); }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (chart.valMin != null) dataMin = chart.valMin;
  else if (dataMin > 0) dataMin = 0;
  if (chart.valMax != null) dataMax = chart.valMax;
  else if (dataMax < 0) dataMax = 0;
  if (dataMax === dataMin) dataMax = dataMin + 1;

  const step  = niceStep(dataMax - dataMin);
  const axMin = chart.valMin ?? niceAxisMin(dataMin, step);
  const axMax = chart.valMax ?? niceAxisMax(dataMax, step);
  const range = axMax - axMin; if (range === 0) return;

  const toY = (v: number) => py0 + ph - ((v - axMin) / range) * ph;
  const toX = (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);

  if (!chart.valAxisHidden) {
    const steps = Math.round((axMax - axMin) / step);
    ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
    ctx.textBaseline = 'middle';
    for (let si = 0; si <= steps; si++) {
      const v = axMin + si * step;
      const gy = toY(v);
      ctx.strokeStyle = v === 0 ? '#aaa' : '#e0e0e0';
      ctx.lineWidth = v === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = '#555'; ctx.textAlign = 'right';
      ctx.fillText(formatChartVal(v), px0 - 4, gy);
    }
  }

  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();

  const markerR = Math.max(3, ph * 0.015);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.setLineDash([]);
    ctx.beginPath();
    let started = false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) { started = false; continue; }
      const px = toX(ci); const py = toY(v);
      if (!started) { ctx.moveTo(px, py); started = true; } else ctx.lineTo(px, py);
    }
    ctx.stroke();
    ctx.fillStyle = color;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) continue;
      ctx.beginPath(); ctx.arc(toX(ci), toY(v), markerR, 0, Math.PI * 2); ctx.fill();
      if (chart.showDataLabels) {
        const lsz = Math.max(7, Math.round(markerR * 1.5));
        ctx.font = `${lsz}px sans-serif`;
        ctx.fillStyle = '#333'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        ctx.fillText(formatChartVal(v), toX(ci), toY(v) - markerR - 1);
        ctx.fillStyle = color;
      }
    }
  }

  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px sans-serif`;
    for (let ci = 0; ci < n; ci += labelInterval) {
      ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), toX(ci), py0 + ph + 3);
    }
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const stacked = chart.chartType === 'stackedArea' || chart.chartType === 'stackedAreaPct';

  const titleH   = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW  = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = { t: titleH + h * 0.04, r: legendW + w * 0.05, b: h * 0.14 + catTitleH, l: w * 0.12 + valTitleW };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    if (stacked) {
      let sum = 0;
      for (const s of chart.series) sum += s.values[ci] ?? 0;
      dataMax = Math.max(dataMax, sum);
    } else {
      for (const s of chart.series) dataMax = Math.max(dataMax, s.values[ci] ?? 0);
    }
  }
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  const step  = niceStep(dataMax);
  const axMax = chart.valMax ?? niceAxisMax(dataMax, step);

  const toX = (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
  const toY = (v: number) => py0 + ph - (v / axMax) * ph;

  if (!chart.valAxisHidden) {
    ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
    ctx.textBaseline = 'middle';
    const steps = Math.round(axMax / step);
    for (let si = 0; si <= steps; si++) {
      const v = si * step; const gy = toY(v);
      ctx.strokeStyle = si === 0 ? '#aaa' : '#e0e0e0';
      ctx.lineWidth = si === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = '#555'; ctx.textAlign = 'right';
      ctx.fillText(formatChartVal(v), px0 - 4, gy);
    }
  }
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();

  const stackBase = stacked ? new Array(n).fill(0) as number[] : null;
  for (let si = chart.series.length - 1; si >= 0; si--) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    const baseY = py0 + ph;

    ctx.beginPath();
    if (stacked && stackBase) {
      for (let ci = 0; ci < n; ci++) {
        const v = (s.values[ci] ?? 0) + stackBase[ci];
        const px = toX(ci); const py = toY(v);
        if (ci === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      for (let ci = n - 1; ci >= 0; ci--) {
        ctx.lineTo(toX(ci), toY(stackBase[ci]));
      }
      for (let ci = 0; ci < n; ci++) stackBase[ci] += s.values[ci] ?? 0;
    } else {
      ctx.moveTo(toX(0), baseY);
      for (let ci = 0; ci < n; ci++) ctx.lineTo(toX(ci), toY(s.values[ci] ?? 0));
      ctx.lineTo(toX(n - 1), baseY);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.6);
    ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([]);
    ctx.stroke();
  }

  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px sans-serif`;
    for (let ci = 0; ci < n; ci += labelInterval) {
      ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), toX(ci), py0 + ph + 3);
    }
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

function renderPieChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, isDoughnut: boolean): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;

  const titleH = chart.title ? Math.max(14, h * 0.06) : 0;
  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const legendW = Math.max(80, w * 0.28);
  const pw = w - legendW; const ph = h - titleH - h * 0.04;
  const cx2 = x + pw / 2; const cy2 = y + titleH + h * 0.04 + ph / 2;
  const outerR = Math.min(pw, ph) * 0.42;
  const innerR = isDoughnut ? outerR * 0.5 : 0;

  let angle = -Math.PI / 2;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const color = pieSliceColor(i, s);
    ctx.beginPath();
    ctx.moveTo(cx2, cy2);
    ctx.arc(cx2, cy2, outerR, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = color; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.stroke();

    if (chart.showDataLabels && slice > 0.15) {
      const midAngle = angle + slice / 2;
      const labelR = outerR * (isDoughnut ? 0.75 : 0.6);
      const lx2 = cx2 + Math.cos(midAngle) * labelR;
      const ly2 = cy2 + Math.sin(midAngle) * labelR;
      const pct2 = Math.round((vals[i] / total) * 100);
      const lsz = Math.max(8, outerR * 0.1);
      ctx.font = `bold ${lsz}px sans-serif`;
      ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(`${pct2}%`, lx2, ly2);
    }

    angle += slice;
  }

  if (isDoughnut) {
    ctx.beginPath(); ctx.arc(cx2, cy2, innerR, 0, Math.PI * 2);
    ctx.fillStyle = '#fff'; ctx.fill();
  }

  const lx = x + pw + 4;
  const fontSize = Math.max(9, Math.min(12, h / (vals.length + 2)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const rowH = fontSize + 4;
  let ry = y + (h - rowH * vals.length) / 2;
  for (let i = 0; i < vals.length; i++) {
    ctx.fillStyle = pieSliceColor(i, s);
    ctx.fillRect(lx, ry, 10, fontSize);
    ctx.fillStyle = '#333'; ctx.textAlign = 'left';
    ctx.fillText((cats[i] ?? `Item ${i + 1}`).toString().slice(0, 18), lx + 14, ry + fontSize / 2);
    ry += rowH;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length > 1 ? Math.max(70, w * 0.2) : 0;
  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const pw = w - legendW; const ph = h - titleH - h * 0.04;
  const cx2 = x + pw / 2; const cy2 = y + titleH + h * 0.04 + ph / 2;
  const rd  = Math.min(pw, ph) * 0.38;

  let dataMax = 0;
  for (const s of chart.series) for (const v of s.values) dataMax = Math.max(dataMax, v ?? 0);
  if (chart.valMax != null) dataMax = chart.valMax;
  if (dataMax === 0) dataMax = 1;
  const step  = niceStep(dataMax);
  const axMax = chart.valMax ?? niceAxisMax(dataMax, step);

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  const rings = Math.round(axMax / step);
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 0.5;
  for (let ri = 1; ri <= rings; ri++) {
    const rr = (ri / rings) * rd;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  }

  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * rd, cy2 + Math.sin(a) * rd); ctx.stroke();
  }

  ctx.font = `${Math.max(8, Math.min(11, rd * 0.2))}px sans-serif`;
  ctx.fillStyle = '#444'; ctx.textBaseline = 'middle';
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const lx = cx2 + Math.cos(a) * (rd + 12);
    const ly = cy2 + Math.sin(a) * (rd + 12);
    ctx.textAlign = Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.fillText((cats[i] ?? '').toString().slice(0, 12), lx, ly);
  }

  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const v = s.values[i] ?? 0;
      const frac = v / axMax;
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rd * frac;
      const py = cy2 + Math.sin(a) * rd * frac;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.25); ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke();
  }

  if (legendW > 0) {
    drawLegend(ctx, chart.series, x + w - legendW + 4, y + titleH + h * 0.04, legendW - 8, ph);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

function renderScatterChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const titleH   = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW  = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = { t: titleH + h * 0.06, r: legendW + w * 0.05, b: h * 0.12 + catTitleH, l: w * 0.12 + valTitleW };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  if (chart.plotAreaBg) {
    ctx.fillStyle = `#${chart.plotAreaBg}`;
    ctx.fillRect(px0, py0, pw, ph);
  }

  const allX: number[] = []; const allY: number[] = [];
  for (const s of chart.series) {
    const cats = s.categories ?? [];
    for (const c of cats) { const v = parseFloat(c); if (!isNaN(v)) allX.push(v); }
    for (const v of s.values) if (v != null) allY.push(v);
  }
  const useIndexX = allX.length === 0;
  if (useIndexX) {
    const maxLen = Math.max(...chart.series.map(s => s.values.length));
    for (let i = 0; i < maxLen; i++) allX.push(i);
  }

  let xMin = Math.min(...allX); let xMax = Math.max(...allX);
  let yMin = Math.min(...allY); let yMax = Math.max(...allY);
  if (xMin === xMax) { xMin -= 1; xMax += 1; }
  if (yMin === yMax) { yMin -= 1; yMax += 1; }
  if (chart.valMin != null) yMin = chart.valMin;
  else if (yMin > 0) yMin = 0;
  if (chart.valMax != null) yMax = chart.valMax;

  const toX = (v: number) => px0 + ((v - xMin) / (xMax - xMin)) * pw;
  const toY = (v: number) => py0 + ph - ((v - yMin) / (yMax - yMin)) * ph;

  if (!chart.valAxisHidden) {
    ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
    const yStep = niceStep(yMax - yMin);
    const ySteps = Math.round((yMax - yMin) / yStep) + 1;
    for (let si = 0; si < ySteps; si++) {
      const v = yMin + si * yStep; if (v > yMax + yStep * 0.01) break;
      const gy = toY(v);
      ctx.strokeStyle = '#e0e0e0'; ctx.lineWidth = 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.fillStyle = '#555'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(formatChartVal(v), px0 - 4, gy);
    }
  }
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();

  const markerR = Math.max(3, ph * 0.015);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    ctx.fillStyle = color;
    const cats = s.categories ?? [];
    for (let ci = 0; ci < s.values.length; ci++) {
      const yv = s.values[ci]; if (yv == null) continue;
      const xv = useIndexX ? ci : parseFloat(cats[ci] ?? '0');
      if (isNaN(xv)) continue;
      ctx.beginPath(); ctx.arc(toX(xv), toY(yv), markerR, 0, Math.PI * 2); ctx.fill();
    }
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

function renderWaterfallChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect): void {
  const { x, y, w, h } = r;
  const padL = w * 0.11;
  const padR = w * 0.04;
  const padT = h * 0.08;
  const padB = h * 0.18;
  const px0 = x + padL;
  const py0 = y + padT;
  const pw  = w - padL - padR;
  const ph  = h - padT - padB;

  const vals = chart.series[0]?.values ?? [];
  const cats = chart.categories;
  const n = cats.length;
  if (n === 0) return;

  const subSet = new Set(chart.subtotalIndices);

  let running = 0;
  const bars: Array<{ start: number; end: number; isSub: boolean; isPos: boolean }> = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i] ?? 0;
    const isSub = i === 0 || subSet.has(i);
    if (isSub) {
      bars.push({ start: 0, end: v, isSub: true, isPos: true });
      running = v;
    } else {
      const start = v >= 0 ? running : running + v;
      const end   = v >= 0 ? running + v : running;
      bars.push({ start, end, isSub: false, isPos: v >= 0 });
      running += v;
    }
  }

  const allEnds = bars.map(b => b.end);
  const allStarts = bars.map(b => b.start);
  const rawMax = Math.max(...allEnds, ...allStarts);
  const rawMin = Math.min(...allStarts, 0);
  const dataRange = rawMax - rawMin;
  if (dataRange <= 0) return;
  const padded = dataRange * 1.1;
  const dataMin = rawMin - dataRange * 0.05;
  const dataMax = dataMin + padded;

  const step = niceStep(padded);
  ctx.save();
  const fontSize = Math.round(h * 0.042);
  ctx.font = `${fontSize}px sans-serif`;
  ctx.strokeStyle = '#e8e8e8';
  ctx.lineWidth = 0.7;
  ctx.fillStyle = '#666';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';
  for (let v = Math.ceil(dataMin / step) * step; v <= dataMax; v += step) {
    const gy = py0 + ph * (1 - (v - dataMin) / padded);
    ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
    ctx.fillText(v.toLocaleString(), px0 - 4, gy);
  }

  ctx.strokeStyle = '#bbb';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(px0, py0);
  ctx.lineTo(px0, py0 + ph);
  ctx.lineTo(px0 + pw, py0 + ph);
  ctx.stroke();

  const colorSub = '#196ECA';
  const colorPos = '#5BA4E6';
  const colorNeg = '#E46970';

  const barW = (pw / n) * 0.55;
  const gapW = pw / n;

  bars.forEach((bar, i) => {
    const bx = px0 + gapW * i + (gapW - barW) / 2;
    const yTop = py0 + ph * (1 - (bar.end - dataMin) / padded);
    const yBot = py0 + ph * (1 - (bar.start - dataMin) / padded);
    const bh = Math.max(1, yBot - yTop);

    if (bar.isSub) {
      ctx.fillStyle = colorSub;
      ctx.fillRect(bx, yTop, barW, bh);
    } else {
      ctx.strokeStyle = bar.isPos ? colorPos : colorNeg;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(bx + 0.75, yTop + 0.75, barW - 1.5, bh - 1.5);
    }

    if (i < n - 1) {
      const nextBx = px0 + gapW * (i + 1) + (gapW - barW) / 2;
      const connY = bar.isPos ? yTop : yBot;
      ctx.strokeStyle = '#ccc';
      ctx.lineWidth = 0.8;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(bx + barW, connY);
      ctx.lineTo(nextBx, connY);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    const rawVal = vals[i] ?? 0;
    const labelText = rawVal < 0
      ? `△ ${Math.abs(rawVal).toLocaleString()}`
      : rawVal.toLocaleString();
    ctx.fillStyle = '#595959';
    ctx.font = `bold ${Math.round(h * 0.044)}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    ctx.fillText(labelText, bx + barW / 2, yTop - 3);
  });

  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#666';
  ctx.font = `${Math.round(h * 0.038)}px sans-serif`;
  const labelY = py0 + ph + 4;
  for (let i = 0; i < n; i++) {
    const ccx = px0 + gapW * i + gapW / 2;
    const lines = cats[i].split(/\s+/);
    lines.forEach((line, li) => ctx.fillText(line, ccx, labelY + li * (fontSize + 2)));
  }

  ctx.restore();
}

// ─── Background frame + dispatcher ──────────────────────────────────────────

/**
 * Render a chart (background frame + dispatch on `chartType`).
 * `rect` is in pixel coordinates on the target canvas.
 */
export function renderChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
): void {
  const { x, y, w, h } = rect;
  // White background + light border (xlsx-style frame). PPTX charts are
  // inside shape boxes with their own backgrounds so the frame is harmless.
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = '#d0d0d0';
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);

  if (chart.series.length === 0) {
    ctx.fillStyle = '#888';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('(no data)', x + w / 2, y + h / 2);
    return;
  }

  switch (chart.chartType) {
    case 'clusteredBar':
    case 'clusteredBarH':
    case 'stackedBar':
    case 'stackedBarH':
    case 'stackedBarPct':
    case 'stackedBarHPct':
      renderBarChart(ctx, chart, rect); break;
    case 'line':
    case 'stackedLine':
    case 'stackedLinePct':
      renderLineChart(ctx, chart, rect); break;
    case 'area':
    case 'stackedArea':
    case 'stackedAreaPct':
      renderAreaChart(ctx, chart, rect); break;
    case 'pie':
      renderPieChart(ctx, chart, rect, false); break;
    case 'doughnut':
      renderPieChart(ctx, chart, rect, true); break;
    case 'radar':
      renderRadarChart(ctx, chart, rect); break;
    case 'scatter':
    case 'bubble':
      renderScatterChart(ctx, chart, rect); break;
    case 'waterfall':
      renderWaterfallChart(ctx, chart, rect); break;
    default:
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`Chart: ${chart.chartType}`, x + w / 2, y + h / 2);
  }
}
