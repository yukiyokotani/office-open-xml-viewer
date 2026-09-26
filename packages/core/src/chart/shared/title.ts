// Classic chart title helpers.
import type { ChartModel, ChartTextRun } from '../../types/chart';
import { cartesianTitleBand, chartTextFontSizePx, resolveManualLayoutRect } from '../layout.js';
import type { ChartTitleBand } from '../layout.js';
import { rawLinkedChartStyleRole } from '../effective-style.js';
import { paintChartLabelBox } from '../label-box.js';
import { chartFontCss, resolveThemeFontRef } from './fonts.js';
import { effectiveLinkedLabelBox } from './style-roles.js';


export interface ResolvedTitlePiece {
  text: string;
  width: number;
  font: string;
  color: string;
}


export interface ResolvedTitleLine {
  pieces: ResolvedTitlePiece[];
  width: number;
  height: number;
}


export function chartTitleRunFont(
  chart: ChartModel,
  run: ChartTextRun,
  fallbackFontSize: number,
): { font: string; fontSize: number; color: string } {
  const titleSizePt = chart.titleFontSizeHpt != null
    && chart.titleFontSizeHpt >= 100
    && chart.titleFontSizeHpt <= 400_000
    ? chart.titleFontSizeHpt / 100
    : 14;
  const effectivePtToPx = fallbackFontSize / titleSizePt;
  const fontSize = chartTextFontSizePx(run.fontSizeHpt, effectivePtToPx) ?? fallbackFontSize;
  const titleFace = resolveThemeFontRef(chart, run.fontFace ?? chart.titleFontFace);
  const face = titleFace ? `"${titleFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
  return {
    font: chartFontCss(
      fontSize,
      face,
      // DrawingML run `b` defaults to false when neither direct text nor an
      // effective numeric/linked title role owns the property. Keep the
      // historical bold fallback only for the public plain-title model, which
      // has no run-level provenance.
      run.bold ?? chart.titleFontBold ?? false,
      run.italic ?? chart.titleFontItalic ?? false,
    ),
    fontSize,
    color: run.colorPaintAuthored === true
      ? run.colorHidden === true || !run.color ? 'transparent' : `#${run.color}`
      : run.color
        ? `#${run.color}`
        : chart.titleFontPaintAuthored === true
          ? chart.titleFontColor ? `#${chart.titleFontColor}` : 'transparent'
          : chart.titleFontColor ? `#${chart.titleFontColor}` : '#333',
  };
}


/** Measure DrawingML title runs against the chart's finite title box. Explicit
 * newlines remain hard breaks; ordinary whitespace is the only automatic wrap
 * opportunity, matching DrawingML's default square text wrapping. */
export function resolveChartTitleLines(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  maxWidth: number,
  fallbackFontSize: number,
): ResolvedTitleLine[] {
  const runs: ChartTextRun[] = chart.titleRichRuns?.length
    ? chart.titleRichRuns
    : chart.title ? [{ text: chart.title }] : [];
  const lines: ResolvedTitleLine[] = [{ pieces: [], width: 0, height: fallbackFontSize }];
  const pushLine = (): ResolvedTitleLine => {
    const line = { pieces: [], width: 0, height: fallbackFontSize } as ResolvedTitleLine;
    lines.push(line);
    return line;
  };
  let line = lines[0];
  for (const run of runs) {
    const style = chartTitleRunFont(chart, run, fallbackFontSize);
    for (const token of run.text.split(/(\n|[\t ]+)/).filter(part => part.length > 0)) {
      if (token === '\n') {
        line = pushLine();
        continue;
      }
      ctx.font = style.font;
      const width = ctx.measureText(token).width;
      const isSpace = /^[\t ]+$/.test(token);
      if (!isSpace && line.pieces.length > 0 && line.width + width > maxWidth) {
        line = pushLine();
      }
      if (isSpace && line.pieces.length === 0) continue;
      line.pieces.push({ text: token, width, font: style.font, color: style.color });
      line.width += width;
      line.height = Math.max(line.height, style.fontSize);
    }
  }
  return lines;
}


export function measuredCartesianTitleBand(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  w: number,
  h: number,
  ptToPx: number,
): ChartTitleBand {
  const base = cartesianTitleBand(chart, h, ptToPx);
  if (base.bandH === 0) return base;
  if (!chart.titleRichRuns?.length) return base;
  const previousFont = ctx.font;
  const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), base.fontPx);
  ctx.font = previousFont;
  const textHeight = lines.reduce((sum, line) => sum + line.height, 0);
  return {
    ...base,
    bandH: base.topPad + textHeight + base.bottomPad,
  };
}


export function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!chart.title) return;
  const titleBox = effectiveLinkedLabelBox(
    chart,
    chart.titleStyle ? { style: chart.titleStyle } : undefined,
    chart.chartStyleRoles?.title,
    rawLinkedChartStyleRole(chart, 'title'),
    true,
  );
  const titlePtToPx = fontSize / Math.max(1, (chart.titleFontSizeHpt ?? 1_400) / 100);
  // Preserve the established single-fillText path for callers/models without
  // formatted DrawingML runs. Besides avoiding unnecessary tokenization, this
  // keeps the public Canvas contract (center-aligned title anchor) unchanged.
  if (!chart.titleRichRuns?.length) {
    const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
    const face = titleFace
      ? `"${titleFace}", Calibri, Arial, sans-serif`
      : 'Calibri, Arial, sans-serif';
    ctx.font = chartFontCss(
      fontSize,
      face,
      chart.titleFontBold ?? true,
      chart.titleFontItalic ?? false,
    );
    ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const measuredWidth = Math.min(w, ctx.measureText(chart.title).width);
    paintChartLabelBox(ctx, titleBox, {
      x: x + (w - measuredWidth) / 2,
      y,
      w: measuredWidth,
      h: fontSize * 1.2,
    }, titlePtToPx);
    ctx.fillText(chart.title, x + w / 2, y);
    return;
  }
  ctx.save();
  const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), fontSize);
  const boxWidth = Math.min(w, Math.max(...lines.map(line => line.width), 0));
  const boxHeight = lines.reduce((sum, line) => sum + line.height, 0);
  paintChartLabelBox(ctx, titleBox, {
    x: x + (w - boxWidth) / 2,
    y,
    w: boxWidth,
    h: boxHeight,
  }, titlePtToPx);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';
  let lineY = y;
  for (const line of lines) {
    let pieceX = x + (w - line.width) / 2;
    for (const piece of line.pieces) {
      ctx.font = piece.font;
      ctx.fillStyle = piece.color;
      ctx.fillText(piece.text, pieceX, lineY);
      pieceX += piece.width;
    }
    lineY += line.height;
  }
  ctx.restore();
}


/** Draw the title at its authored manual-layout position. Office ignores w/h
 * for title descendants and fits the box to text (MS-OI29500 §2.1.1573), while
 * x/y still use the shared factor/edge rules. */
export function drawChartTitleForLayout(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number, y: number, w: number, h: number,
  defaultY: number,
  fontSize: number,
): void {
  if (!chart.title) return;
  const ml = chart.titleManualLayout;
  if (ml) {
    const titleFace = resolveThemeFontRef(chart, chart.titleFontFace);
    const face = titleFace ? `"${titleFace}", Calibri, Arial, sans-serif` : 'Calibri, Arial, sans-serif';
    ctx.font = chartFontCss(
      fontSize,
      face,
      chart.titleFontBold ?? true,
      chart.titleFontItalic ?? false,
    );
    const lines = resolveChartTitleLines(ctx, chart, Math.max(1, w), fontSize);
    const autoWidth = Math.min(w, Math.max(...lines.map(line => line.width), 0));
    const automatic = {
      x: x + (w - autoWidth) / 2,
      y: defaultY,
      w: autoWidth,
      h: fontSize,
    };
    const resolved = resolveManualLayoutRect(
      { ...ml, w: undefined, h: undefined },
      { x, y, w, h },
      automatic,
    );
    if (resolved) {
      drawChartTitle(ctx, chart, resolved.x, resolved.y, resolved.w, fontSize);
      return;
    }
  }
  drawChartTitle(ctx, chart, x, defaultY, w, fontSize);
}
