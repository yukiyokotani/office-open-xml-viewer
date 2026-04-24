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
  // Matches Excel's default `<c:valAx><c:numFmt formatCode="General">` which
  // shows raw numbers — no "k"/"M" abbreviation.
  if (Number.isInteger(v)) return String(v);
  // Trim trailing zeros on decimals (so 0.50 → "0.5") but cap at 6 digits.
  return v.toFixed(6).replace(/\.?0+$/, '');
}

/**
 * Format a chart value with an Excel number-format code. Honors ECMA-376
 * §18.8.30 section syntax (positive;negative;zero;text), common literal
 * escapes (`"..."`, `\x`, `_x` → space), and numeric patterns built from
 * `#`, `0`, `.`, `,`. Unknown tokens are emitted verbatim so currency
 * symbols like `¥` or `$` keep working even when the workbook stored them
 * unquoted. Returns the default `formatChartVal` output when `code` is null
 * or an empty section tells the caller to hide the value.
 */
function formatChartValWithCode(v: number, code: string | null | undefined): string {
  if (!code) return formatChartVal(v);
  const sections = splitFormatSections(code);
  // Section selection per §18.8.30: positive;negative;zero;text. When the
  // negative section is omitted a negative number is formatted with the
  // positive section and a leading minus, which the caller must prepend.
  let section: string;
  if (v > 0) section = sections[0] ?? code;
  else if (v < 0) section = sections[1] ?? sections[0] ?? code;
  else section = sections[2] ?? sections[0] ?? code;
  if (section === '') return '';
  // Negative-without-explicit-section: format absolute value with positive
  // section and prepend '-' unless the section itself already begins with a
  // literal minus.
  const needsLeadingMinus = v < 0 && sections.length < 2;
  const abs = Math.abs(v);
  return (needsLeadingMinus ? '-' : '') + applyChartNumberSection(abs, section);
}

/**
 * Split a format code on unescaped semicolons. Quotes, `[...]` metadata, and
 * `\;` are treated as opaque so `"a;b"` and `\;` stay in a single section.
 */
function splitFormatSections(code: string): string[] {
  const out: string[] = [];
  let buf = '';
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '\\' && i + 1 < code.length) { buf += c + code[i + 1]; i++; continue; }
    if (c === '"') {
      buf += c;
      i++;
      while (i < code.length && code[i] !== '"') { buf += code[i]; i++; }
      if (i < code.length) buf += code[i];
      continue;
    }
    if (c === '[') {
      buf += c;
      i++;
      while (i < code.length && code[i] !== ']') { buf += code[i]; i++; }
      if (i < code.length) buf += code[i];
      continue;
    }
    if (c === ';') { out.push(buf); buf = ''; continue; }
    buf += c;
  }
  out.push(buf);
  return out;
}

function applyChartNumberSection(abs: number, section: string): string {
  // Tokenize the section, separating numeric-pattern runs (`#`, `0`, `.`,
  // `,`, `?`) from literal runs so percent / decimal handling runs once.
  type Tok = { kind: 'lit' | 'num'; text: string };
  const toks: Tok[] = [];
  let i = 0;
  let pushedNum = false;
  let percent = false;
  while (i < section.length) {
    const c = section[i];
    if (c === '"') {
      i++;
      let s = '';
      while (i < section.length && section[i] !== '"') { s += section[i]; i++; }
      if (i < section.length) i++;
      toks.push({ kind: 'lit', text: s });
      continue;
    }
    if (c === '\\' && i + 1 < section.length) {
      toks.push({ kind: 'lit', text: section[i + 1] });
      i += 2;
      continue;
    }
    if (c === '_' && i + 1 < section.length) {
      // `_x` pads a width of x — render as a single space, matching Excel
      // alignment padding without caring about exact glyph metrics.
      toks.push({ kind: 'lit', text: ' ' });
      i += 2;
      continue;
    }
    if (c === '*' && i + 1 < section.length) {
      // `*x` fills the remaining column width with x; we can't know the
      // column width at this layer so we drop it.
      i += 2;
      continue;
    }
    if (c === '[') {
      i++;
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
      continue;
    }
    if (c === '%') { percent = true; toks.push({ kind: 'lit', text: '%' }); i++; continue; }
    if (c === '#' || c === '0' || c === '.' || c === ',' || c === '?') {
      let run = '';
      while (
        i < section.length &&
        (section[i] === '#' || section[i] === '0' || section[i] === '.' ||
         section[i] === ',' || section[i] === '?')
      ) { run += section[i]; i++; }
      toks.push({ kind: 'num', text: run });
      pushedNum = true;
      continue;
    }
    // Everything else (currency symbols like ¥, $, parens, spaces) is literal.
    toks.push({ kind: 'lit', text: c });
    i++;
  }
  if (!pushedNum) {
    // No numeric pattern at all — section is purely literal (e.g. `"N/A"`).
    return toks.map(t => t.text).join('');
  }
  const value = percent ? abs * 100 : abs;
  // Merge numeric tokens into one pattern — Excel treats `#,##0.00` as a
  // single pattern even when flanked by literals. We keep the literal tokens
  // where they are and replace the first num token with the formatted number,
  // dropping subsequent num tokens (they're all part of the same pattern).
  let pattern = '';
  for (const t of toks) if (t.kind === 'num') pattern += t.text;
  const formatted = formatNumericPattern(value, pattern);
  let seenNum = false;
  return toks.map(t => {
    if (t.kind === 'lit') return t.text;
    if (seenNum) return '';
    seenNum = true;
    return formatted;
  }).join('');
}

function formatNumericPattern(value: number, pattern: string): string {
  // Detect thousands separator (a `,` between digit placeholders) and the
  // number of decimal places (digit chars after `.`).
  let dotIdx = pattern.indexOf('.');
  const intPart = dotIdx >= 0 ? pattern.slice(0, dotIdx) : pattern;
  const fracPart = dotIdx >= 0 ? pattern.slice(dotIdx + 1) : '';
  const thousands = /,/.test(intPart);
  const fracDigits = (fracPart.match(/[#0?]/g) ?? []).length;
  // Minimum integer digits = count of `0` in integer part.
  const minIntDigits = (intPart.replace(/,/g, '').match(/0/g) ?? []).length;
  const rounded = value.toFixed(fracDigits);
  const [ints, fracs = ''] = rounded.split('.');
  const paddedInts = ints.padStart(minIntDigits, '0');
  const withSeparators = thousands ? paddedInts.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : paddedInts;
  if (fracDigits === 0) return withSeparators;
  return `${withSeparators}.${fracs.padEnd(fracDigits, '0')}`;
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
  orient: 'vertical' | 'horizontal' = 'vertical',
): void {
  const sw = 10; const gap = 4;
  if (orient === 'horizontal') {
    // Excel lays a bottom/top legend as a single horizontal row, centered.
    const fontSize = Math.max(9, Math.min(12, lh * 0.7));
    ctx.font = `${fontSize}px sans-serif`;
    ctx.textBaseline = 'middle';
    const itemGap = 12;
    const labels = series.map((s, i) => s.name || `Series ${i + 1}`);
    const itemWidths = labels.map((l) => sw + gap + ctx.measureText(l.slice(0, 30)).width);
    const total = itemWidths.reduce((a, b) => a + b, 0) + itemGap * Math.max(0, series.length - 1);
    let rx = lx + (lw - total) / 2;
    const ry = ly + lh / 2;
    for (let i = 0; i < series.length; i++) {
      ctx.fillStyle = chartColor(i, series[i]);
      ctx.fillRect(rx, ry - fontSize / 2, sw, fontSize);
      ctx.fillStyle = '#333'; ctx.textAlign = 'left';
      ctx.fillText(labels[i].slice(0, 30), rx + sw + gap, ry);
      rx += itemWidths[i] + itemGap;
    }
    return;
  }
  const fontSize = Math.max(9, Math.min(12, lh / (series.length + 1)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
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

type LegendSide = 'r' | 'l' | 't' | 'b';
interface LegendLayout {
  side: LegendSide;
  /** Reserved plot-area width (>0 when side = l or r). */
  reserveW: number;
  /** Reserved plot-area height (>0 when side = t or b). */
  reserveH: number;
}

/** Resolve legend placement from `<c:legendPos>`. Returns null when hidden. */
function legendLayout(chart: ChartModel, w: number, h: number): LegendLayout | null {
  if (!chart.showLegend) return null;
  const pos = chart.legendPos ?? 'r';
  const side: LegendSide = pos === 'l' ? 'l' : pos === 't' ? 't' : pos === 'b' ? 'b' : 'r';
  if (side === 'r' || side === 'l') {
    return { side, reserveW: Math.max(80, w * 0.22), reserveH: 0 };
  }
  // Excel's top/bottom legend is a single-row strip; reserve ~8% of height.
  return { side, reserveW: 0, reserveH: Math.max(18, h * 0.08) };
}

/** Draw a legend in the band reserved by {@link legendLayout}. */
function drawLegendForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  leg: LegendLayout | null,
  x: number, y: number, w: number, h: number,
  px0: number, py0: number, pw: number, ph: number,
  topBand: number,
): void {
  if (!leg) return;
  // `<c:legend><c:manualLayout>` (§21.2.2.31) wins over the default side-based
  // rectangle. We honor the `edge` placement mode — fractions are measured
  // from the top-left of the chart space — which matches what Excel's built-in
  // templates emit. `factor` mode (offset from default) is rarer; fall back to
  // the reserved band in that case rather than guess.
  const ml = chart.legendManualLayout;
  if (ml && ml.xMode === 'edge' && ml.yMode === 'edge' && ml.w > 0 && ml.h > 0) {
    const lx = x + ml.x * w;
    const ly = y + ml.y * h;
    const lw = ml.w * w;
    const lh = ml.h * h;
    // Legend is always a horizontal strip when placed on top/bottom; vertical
    // when on left/right. A manual box wider than tall implies horizontal —
    // matches Excel's one-row legend rendering for top/bottom manual layouts.
    const orient = lw >= lh ? 'horizontal' : 'vertical';
    drawLegend(ctx, chart.series, lx, ly, lw, lh, orient);
    return;
  }
  switch (leg.side) {
    case 'r':
      drawLegend(ctx, chart.series, x + w - leg.reserveW + 4, py0, leg.reserveW - 8, ph);
      break;
    case 'l':
      drawLegend(ctx, chart.series, x + 4, py0, leg.reserveW - 8, ph);
      break;
    case 't':
      drawLegend(ctx, chart.series, px0, y + topBand, pw, leg.reserveH, 'horizontal');
      break;
    case 'b':
      drawLegend(ctx, chart.series, px0, y + h - leg.reserveH, pw, leg.reserveH, 'horizontal');
      break;
  }
}

function drawAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string,
  axis: 'val' | 'cat',
  anchorXOrY: number,
  perpendicular: number,
): void {
  if (mode === 'none' || !mode) return;
  const len = 4;
  const prev = ctx.strokeStyle;
  ctx.strokeStyle = '#888';
  ctx.lineWidth = 1;
  ctx.beginPath();
  if (axis === 'val') {
    // val axis is vertical (x = anchor, y varies). Ticks extend horizontally.
    const x0 = anchorXOrY;
    const y = perpendicular;
    const outer = mode === 'out' || mode === 'cross' ? -len : 0;
    const inner = mode === 'in' || mode === 'cross' ? len : 0;
    ctx.moveTo(x0 + outer, y);
    ctx.lineTo(x0 + inner, y);
  } else {
    // cat axis is horizontal (y = anchor, x varies). Ticks extend vertically.
    const y0 = anchorXOrY;
    const xc = perpendicular;
    const outer = mode === 'out' || mode === 'cross' ? len : 0;
    const inner = mode === 'in' || mode === 'cross' ? -len : 0;
    ctx.moveTo(xc, y0 + outer);
    ctx.lineTo(xc, y0 + inner);
  }
  ctx.stroke();
  ctx.strokeStyle = prev;
}

function chartTitleFontPx(chart: ChartModel, h: number, ptToPx: number): number {
  // Honor the XML-specified title font size (hundredths of a point) when
  // present. ptToPx is the pixels-per-point at the current slide scale, so
  // a 16pt title renders at the same proportional size as PowerPoint.
  if (chart.titleFontSizeHpt) return (chart.titleFontSizeHpt / 100) * ptToPx;
  return Math.max(10, h * 0.085);
}

/** Resolve an axis label font size (px) from <c:txPr> hpt or a proportional
 *  fallback. ptToPx comes from the host renderer (EMU/px scale at display). */
function axisLabelPx(sizeHpt: number | null | undefined, h: number, ptToPx: number): number {
  if (sizeHpt) return (sizeHpt / 100) * ptToPx;
  return Math.max(8, h * 0.045);
}

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  const face = chart.titleFontFace ? `"${chart.titleFontFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
  ctx.font = `bold ${fontSize}px ${face}`;
  ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(chart.title, x + w / 2, y);
}

// ─── Category helper ────────────────────────────────────────────────────────

function chartCategories(chart: ChartModel): string[] {
  if (chart.categories.length > 0) return chart.categories;
  const first = chart.series[0];
  if (first?.categories && first.categories.length > 0) return first.categories;
  // ECMA-376 §21.2.2.24 — when <c:cat> is absent the category axis uses
  // integer values starting at 1. Fall back to the longest series so the
  // chart still renders instead of bailing out at n === 0.
  let n = 0;
  for (const s of chart.series) if (s.values.length > n) n = s.values.length;
  return n > 0 ? Array.from({ length: n }, (_, i) => String(i + 1)) : [];
}

/**
 * Draw a bar data label with the ECMA-376 §21.2.2.16 `dLblPos` semantics.
 *
 * For a vertical bar the coordinates describe the rectangle top-left + width +
 * height; for a horizontal bar they describe the bar's left-edge `bx`, top `by`,
 * length `barL`, and thickness `barW`. When `position` is "inBase" / "inEnd" /
 * "ctr" the label sits inside the bar; "outEnd" (default for clustered bars)
 * nudges the text just past the far edge. An explicit `color` overrides the
 * default dark label fill — Excel's workbook typically pairs "inBase" with a
 * white text color so labels stay readable against the bar fill.
 */
function drawBarDataLabel(
  ctx: CanvasRenderingContext2D,
  text: string,
  bx: number, by: number, barL: number, barW: number,
  orient: 'vertical' | 'horizontal',
  position: string | null,
  color: string | null,
): void {
  const pos = (position ?? 'outEnd');
  const fill = color ? `#${color}` : '#333';
  ctx.fillStyle = fill;
  if (orient === 'vertical') {
    // bx/by = top-left of bar rect (bar grows upward from by+barL toward by).
    // barL here is bar height (pixels) and barW is bar width.
    const cx = bx + barW / 2;
    if (pos === 'inBase') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ctx.fillText(text, cx, by + barL - 2);
    } else if (pos === 'inEnd') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ctx.fillText(text, cx, by + 2);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, cx, by + barL / 2);
    } else {
      // outEnd / default: just above the bar's top edge (by).
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      ctx.fillText(text, cx, by - 1);
    }
  } else {
    // Horizontal: bar grows to the right from bx.
    const cy = by + barW / 2;
    if (pos === 'inBase') {
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + 4, cy);
    } else if (pos === 'inEnd') {
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL - 4, cy);
    } else if (pos === 'ctr') {
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL / 2, cy);
    } else {
      // outEnd / default: just past the bar's right edge.
      ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
      ctx.fillText(text, bx + barL + 2, cy);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

function renderBarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const barSeries  = chart.series.filter(s => s.seriesType !== 'line');
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;

  // Honor the XML-specified title font size when present; otherwise fall back
  // to the proportional heuristic. Reserve the title band based on the actual
  // drawn height so the plot shrinks to avoid overlap.
  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.02 : 0;
  const titleBottomPad = chart.title ? h * 0.025 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = {
    t: titleH + legTopH + h * 0.02,
    r: legRightW + w * 0.03,
    b: h * 0.14 + catTitleH + legBottomH,
    l: w * 0.12 + valTitleW + legLeftW,
  };
  if (isH) {
    // With the category axis hidden (`c:catAx/c:delete val="1"`) there are no
    // category tick labels to reserve room for — tighten the left margin so
    // the bars can extend to the chart edge, matching Excel's rendering.
    pad.l = (chart.catAxisHidden ? w * 0.03 : w * 0.22) + valTitleW + legLeftW;
    pad.b = (chart.valAxisHidden ? h * 0.02 : h * 0.08) + catTitleH + legBottomH;
  }

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

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
      const label = pct
        ? `${Math.round(val)}%`
        : formatChartValWithCode(val, chart.valAxisFormatCode);
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

  // Bar cluster geometry — ECMA-376 §21.2.2.13 (gapWidth = % of bar width
  // between categories, default 150) and §21.2.2.25 (overlap = signed % of
  // bar width within a cluster, default 0). Within a cluster the pitch
  // between consecutive bars is `barW * (1 - overlap/100)`, so with N series:
  //   clusterWidth = barW + (N - 1) * barW * (1 - overlap/100)
  //   catGap       = clusterWidth + barW * gapWidth/100
  //                = barW * (1 + (N-1) * (1 - overlap/100) + gapWidth/100)
  // Solving for barW gives the formula below. Stacked charts render one bar
  // per category so we treat them as N=1 and overlap=0.
  const catGap = !isH ? pw / n : ph / n;
  const nSeriesEffective = stacked ? 1 : Math.max(1, barSeries.length);
  const overlapPct  = stacked ? 0 : (chart.barOverlap ?? 0);
  const gapWidthPct = chart.barGapWidth ?? 150;
  const denom = 1 + (nSeriesEffective - 1) * (1 - overlapPct / 100) + gapWidthPct / 100;
  const barW  = catGap / denom;
  // Pitch between bars within a cluster (not the gap — the left-edge to
  // left-edge distance). Kept named `clusterGap` for continuity with the
  // prior implementation, which also used it as a pitch.
  const clusterGap = stacked ? 0 : barW * (1 - overlapPct / 100);
  const clusterWidth = barW + (nSeriesEffective - 1) * clusterGap;
  // Center the cluster inside the category slot.
  const catStart   = (catGap - clusterWidth) / 2;

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
          const lsz = Math.max(7, Math.min(11, barW * 0.6));
          ctx.font = `bold ${lsz}px sans-serif`;
          const text = pct
            ? `${Math.round(val)}%`
            : formatChartValWithCode(
                val,
                chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
              );
          drawBarDataLabel(
            ctx, text,
            bx, by, barW, barH,
            'vertical',
            chart.dataLabelPosition ?? null,
            chart.dataLabelFontColor ?? null,
          );
        }
      } else {
        // Excel renders horizontal clustered bars with series 0 at the BOTTOM
        // of each category cluster (so the legend's top entry matches the bar
        // at the top of the plot). Reverse the per-series offset so `order=0`
        // ends up at the bottom; stacked horizontal bars use a single anchor.
        const siVisual = stacked ? si : (barSeries.length - 1 - si);
        const by = stacked
          ? py0 + (n - 1 - ci) * catGap + catStart
          : py0 + (n - 1 - ci) * catGap + catStart + siVisual * clusterGap;
        const barL = (val / axMax) * pw;
        const bx   = stacked ? px0 + (stackOffset / axMax) * pw : px0;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barL, barW);
        if (chart.showDataLabels && val > 0) {
          const lsz = Math.max(7, Math.min(11, barW * 0.6));
          ctx.font = `bold ${lsz}px sans-serif`;
          const text = pct
            ? `${Math.round(val)}%`
            : formatChartValWithCode(
                val,
                chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
              );
          drawBarDataLabel(
            ctx, text,
            bx, by, barL, barW,
            'horizontal',
            chart.dataLabelPosition ?? null,
            chart.dataLabelFontColor ?? null,
          );
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
      if (s.showMarker !== false) {
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
  }

  // Horizontal clustered bars: Excel mirrors the series order between the
  // plot and the legend so the legend's first entry matches the top bar. We
  // already flipped the bar rendering; reverse the legend series too.
  const legendChart = isH && !stacked
    ? { ...chart, series: [...chart.series].reverse() }
    : chart;
  drawLegendForLayout(ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function renderLineChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  // PowerPoint's auto-layout reserves a title band with air above and below
  // the text; pinning the title to y+0 and the plot to y+titleFontPx+2 is too
  // tight. Use proportional pads so scaling preserves the same feel.
  const titleTopPad    = chart.title ? h * 0.045 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const axisFontSz = Math.max(catAxFontPx, valAxFontPx);
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  // Pad based on actual label metrics rather than magic percents so an explicit
  // <c:txPr sz="1000"> (10pt) correctly compresses the plot area.
  const pad = {
    t: titleH + legTopH + valAxFontPx / 2 + 2,
    r: legRightW + w * 0.05,
    b: catAxFontPx + 12 + catTitleH + legBottomH,
    l: valAxFontPx * 2.2 + 10 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

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
  // crossBetween="between" (default) insets the first/last category by half a
  // step so points aren't flush against the axes. "midCat" anchors them.
  const between = chart.catAxisCrossBetween !== 'midCat';
  const toX = between
    ? (i: number) => px0 + ((i + 0.5) / n) * pw
    : (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);

  if (!chart.valAxisHidden) {
    const steps = Math.round((axMax - axMin) / step);
    ctx.font = `${valAxFontPx}px sans-serif`;
    ctx.textBaseline = 'middle';
    for (let si = 0; si <= steps; si++) {
      const v = axMin + si * step;
      const gy = toY(v);
      ctx.strokeStyle = v === 0 ? '#aaa' : '#e0e0e0';
      ctx.lineWidth = v === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy);
      ctx.fillStyle = '#555'; ctx.textAlign = 'right';
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 6, gy);
    }
  }

  // Axis lines: bottom (category) + left (value). Both default to visible
  // unless hidden explicitly.
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  if (!chart.valAxisHidden) {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  // Line width and marker size come from OOXML in points (<a:ln w=EMU> /
  // <c:marker><c:size val=pt>). We don't parse per-series overrides yet so
  // use the PowerPoint defaults (2.25pt line, 5pt marker diameter) scaled to
  // the current slide pt-per-px so both shrink with the viewport.
  const lineWidthPx = Math.max(1, 2.25 * ptToPx);
  const markerR = Math.max(2, 2.5 * ptToPx);
  const dataLabelPx = axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    ctx.strokeStyle = color; ctx.lineWidth = lineWidthPx; ctx.setLineDash([]);
    ctx.beginPath();
    let started = false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) { started = false; continue; }
      const px = toX(ci); const py = toY(v);
      if (!started) { ctx.moveTo(px, py); started = true; } else ctx.lineTo(px, py);
    }
    ctx.stroke();
    ctx.fillStyle = color;
    // ECMA-376 §21.2.2.32 — when the series resolves to no marker, skip the
    // data-point dots but keep data labels (which pin to each raw value, not
    // to the marker).
    const drawMarkers = s.showMarker !== false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) continue;
      if (drawMarkers) {
        ctx.beginPath(); ctx.arc(toX(ci), toY(v), markerR, 0, Math.PI * 2); ctx.fill();
      }
      if (chart.showDataLabels) {
        ctx.font = `${dataLabelPx}px sans-serif`;
        ctx.fillStyle = '#333'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        const labelOffset = drawMarkers ? markerR + 1 : 2;
        ctx.fillText(formatChartVal(v), toX(ci), toY(v) - labelOffset);
        ctx.fillStyle = color;
      }
    }
  }

  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.ceil(n / 8));
    ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = `${catAxFontPx}px sans-serif`;
    for (let ci = 0; ci < n; ci += labelInterval) {
      const tx = toX(ci);
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx);
      ctx.fillStyle = '#555';
      ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), tx, py0 + ph + 5);
    }
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const stacked = chart.chartType === 'stackedArea' || chart.chartType === 'stackedAreaPct';

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = {
    t: titleH + legTopH + h * 0.02,
    r: legRightW + w * 0.05,
    b: h * 0.14 + catTitleH + legBottomH,
    l: w * 0.12 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

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
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 4, gy);
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

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
  if (chart.catAxisTitle) drawAxisTitle(ctx, chart.catAxisTitle, px0, py0, pw, ph, 'cat', axisFontSz);
  if (chart.valAxisTitle) drawAxisTitle(ctx, chart.valAxisTitle, px0, py0, pw, ph, 'val', axisFontSz);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

function renderPieChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, isDoughnut: boolean, ptToPx: number): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  // Pie legend labels categories (one row per slice) so reserve a bit more
  // than the default 22% when placed on the side.
  const pieLeg: LegendLayout | null = chart.showLegend
    ? (() => {
        const pos = chart.legendPos ?? 'r';
        const side: LegendSide = pos === 'l' ? 'l' : pos === 't' ? 't' : pos === 'b' ? 'b' : 'r';
        if (side === 'r' || side === 'l') {
          return { side, reserveW: Math.max(80, w * 0.28), reserveH: 0 };
        }
        return { side, reserveW: 0, reserveH: Math.max(18, h * 0.08) };
      })()
    : null;
  const legRightW  = pieLeg?.side === 'r' ? pieLeg.reserveW : 0;
  const legLeftW   = pieLeg?.side === 'l' ? pieLeg.reserveW : 0;
  const legTopH    = pieLeg?.side === 't' ? pieLeg.reserveH : 0;
  const legBottomH = pieLeg?.side === 'b' ? pieLeg.reserveH : 0;

  const pw = w - legRightW - legLeftW;
  const ph = h - titleH - legTopH - legBottomH - h * 0.02;
  const cx2 = x + legLeftW + pw / 2;
  const cy2 = y + titleH + legTopH + h * 0.02 + ph / 2;
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

  if (pieLeg) {
    // Pie legend is category-driven; build a pseudo-series array whose per-
    // index palette matches the pie's slice colors so the swatches line up.
    const legendSeries: ChartSeries[] = vals.map((_, i) => ({
      name: (cats[i] ?? `Item ${i + 1}`).toString(),
      color: s.dataPointColors?.[i] ?? s.color ?? CHART_PALETTE[i % CHART_PALETTE.length],
      values: [],
    }));
    const plotLeft = cx2 - pw / 2;
    drawLegendForLayout(
      ctx, { ...chart, series: legendSeries } as ChartModel, pieLeg,
      x, y, w, h, plotLeft, cy2 - ph / 2, pw, ph, titleH + 2,
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH  = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

  const pw = w - legRightW - legLeftW;
  const ph = h - titleH - legTopH - legBottomH - h * 0.02;
  const cx2 = x + legLeftW + pw / 2;
  const cy2 = y + titleH + legTopH + h * 0.02 + ph / 2;
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

  // Radial tick labels on the top (12 o'clock) spoke — Excel places the value
  // axis there for radar charts. Respect <c:valAx><c:delete val="1"/> when the
  // caller hides the axis, and skip the 0-label at the center to avoid
  // overlapping the origin point.
  if (!chart.valAxisHidden) {
    const valAxPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
    ctx.font = `${valAxPx}px sans-serif`;
    ctx.fillStyle = '#555';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (let ri = 1; ri <= rings; ri++) {
      const v = (ri / rings) * axMax;
      const rr = (ri / rings) * rd;
      ctx.fillText(formatChartVal(v), cx2 - 3, cy2 - rr);
    }
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

  drawLegendForLayout(
    ctx, chart, leg,
    x, y, w, h,
    cx2 - pw / 2, cy2 - ph / 2, pw, ph, titleH + 2,
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

function renderScatterChart(ctx: CanvasRenderingContext2D, chart: ChartModel, r: ChartRect, ptToPx: number): void {
  const { x, y, w, h } = r;
  const titleFontPx = chart.title ? chartTitleFontPx(chart, h, ptToPx) : 0;
  const titleTopPad    = chart.title ? h * 0.035 : 0;
  const titleBottomPad = chart.title ? h * 0.035 : 0;
  const titleH   = chart.title ? titleFontPx + titleTopPad + titleBottomPad : 0;
  const leg = legendLayout(chart, w, h);
  const legRightW  = leg?.side === 'r' ? leg.reserveW : 0;
  const legLeftW   = leg?.side === 'l' ? leg.reserveW : 0;
  const legTopH    = leg?.side === 't' ? leg.reserveH : 0;
  const legBottomH = leg?.side === 'b' ? leg.reserveH : 0;
  const axisFontSz = Math.max(8, Math.min(10, h * 0.045));
  const catTitleH  = chart.catAxisTitle ? axisFontSz + 4 : 0;
  const valTitleW  = chart.valAxisTitle ? axisFontSz + 4 : 0;
  const pad = {
    t: titleH + legTopH + h * 0.02,
    r: legRightW + w * 0.05,
    b: h * 0.12 + catTitleH + legBottomH,
    l: w * 0.12 + valTitleW + legLeftW,
  };

  drawChartTitle(ctx, chart, x, y + titleTopPad, w, titleFontPx);

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
      ctx.fillText(formatChartValWithCode(v, chart.valAxisFormatCode), px0 - 4, gy);
    }
  }
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();

  const markerR = Math.max(3, ph * 0.015);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    if (s.showMarker === false) continue;
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

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2);
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
  /**
   * Pixels per point at the caller's current display scale. For PPTX at
   * 960px/12192000EMU the value is ~1.05; xlsx's sheet view renders at
   * device-px where 1pt≈1.333. Used to size title/axis labels whose
   * XML-specified sizes are in OOXML hundredths of a point.
   */
  ptToPx: number = 1.333,
): void {
  const { x, y, w, h } = rect;
  // Only fill the outer chartSpace when chartBg is set; a null means noFill
  // (transparent) per OOXML, so the underlying slide/sheet shows through.
  if (chart.chartBg) {
    ctx.fillStyle = `#${chart.chartBg}`;
    ctx.fillRect(x, y, w, h);
  }

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
      renderBarChart(ctx, chart, rect, ptToPx); break;
    case 'line':
    case 'stackedLine':
    case 'stackedLinePct':
      renderLineChart(ctx, chart, rect, ptToPx); break;
    case 'area':
    case 'stackedArea':
    case 'stackedAreaPct':
      renderAreaChart(ctx, chart, rect, ptToPx); break;
    case 'pie':
      renderPieChart(ctx, chart, rect, false, ptToPx); break;
    case 'doughnut':
      renderPieChart(ctx, chart, rect, true, ptToPx); break;
    case 'radar':
      renderRadarChart(ctx, chart, rect, ptToPx); break;
    case 'scatter':
    case 'bubble':
      renderScatterChart(ctx, chart, rect, ptToPx); break;
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
