// Classic chart trendline helpers.
import type { ChartModel, ChartRect, ChartSeries, ChartTrendline } from '../../types/chart';
import { formatChartVal, formatChartValWithCode } from '../chart-number-format.js';
import { fitTrendline, linearTrendlineStats } from '../axis-scale.js';
import { chartTextFontSizePx } from '../layout.js';
import { paintRichDataLabelBlock, resolveRichDataLabelBlock } from '../rich-data-label.js';
import { dataLabelInsets, fitStyledDataLabelLines, rotatedDataLabelSize } from '../data-label-style.js';
import type { DataLabelTextStyle } from '../data-label-style.js';
import { placeTrendlineLabel } from '../trendline-label.js';
import { paintChartLabelBox } from '../label-box.js';
import { elideToWidth } from '../text-elide.js';
import { axisLineWidthPx } from '../axis-style.js';
import { chartVariesColorsByPoint, legendIsCategoryDriven } from '../legend-entry-plan.js';
import { chartFontCss, chartFontFamily } from './fonts.js';
import { dashPatternForPreset } from './geometry.js';


export interface TrendlineLabelContext {
  chart: ChartModel;
  chartRect: ChartRect;
  plotRect: ChartRect;
  clipLineToPlot?: boolean;
  automaticAnchor?: { x: number; y: number };
  shapeRotationDeg?: number;
}


export function compactTrendlineNumber(value: number, formatCode?: string | null): string {
  if (formatCode && formatCode.trim().toLowerCase() !== 'general') {
    return formatChartValWithCode(value, formatCode);
  }
  return formatChartVal(Number(value.toPrecision(6)));
}


export function generatedTrendlineLabel(
  tl: ChartTrendline,
  stats: ReturnType<typeof linearTrendlineStats>,
  sourceFormatCode?: string | null,
): string[] {
  if (tl.labelText) return tl.labelText.split(/\r?\n/);
  if (!stats) return [];
  const formatCode = tl.labelFormatSourceLinked === true
    ? sourceFormatCode
    : tl.labelFormatCode;
  const lines: string[] = [];
  if (tl.dispEq) {
    const sign = stats.intercept < 0 ? '−' : '+';
    lines.push(
      `y = ${compactTrendlineNumber(stats.slope, formatCode)}x ${sign} ${compactTrendlineNumber(Math.abs(stats.intercept), formatCode)}`,
    );
  }
  if (tl.dispRSqr) lines.push(`R² = ${compactTrendlineNumber(stats.rSquared, formatCode)}`);
  return lines;
}


export function drawTrendlineLabel(
  ctx: CanvasRenderingContext2D,
  tl: ChartTrendline,
  stats: ReturnType<typeof linearTrendlineStats>,
  ptToPx: number,
  labelContext?: TrendlineLabelContext,
  sourceFormatCode?: string | null,
): void {
  if (!labelContext) return;
  const lines = generatedTrendlineLabel(tl, stats, sourceFormatCode);
  if (lines.length === 0) return;
  const { chart, chartRect, plotRect } = labelContext;
  const fontPx = chartTextFontSizePx(tl.labelFontSizeHpt, ptToPx)
    ?? chartTextFontSizePx(chart.dataLabelFontSizeHpt, ptToPx)
    ?? 10 * ptToPx;
  const face = chartFontFamily(chart, tl.labelFontFace ?? chart.dataLabelFontFace, 'minor');
  const bold = tl.labelFontBold ?? chart.dataLabelFontBold ?? false;
  const italic = tl.labelFontItalic ?? false;
  ctx.font = chartFontCss(fontPx, face, bold, italic);
  const lineHeight = fontPx * 1.2;
  const color = tl.labelFontColor ?? chart.dataLabelFontColor;
  const rich = tl.labelRichRuns?.length
      ? resolveRichDataLabelBlock(ctx, {
        runs: tl.labelRichRuns,
        ptToPx,
        fontFamily: face,
        fallbackBold: bold,
        fallbackItalic: italic,
        fallbackBaseline: tl.labelFontBaseline ?? undefined,
        fallbackColorHidden: tl.labelFontPaintAuthored === true
          && (tl.labelFontHidden === true || tl.labelFontColor == null),
        fontFamilyForFace: runFace => chartFontFamily(chart, runFace, 'minor'),
      }, fontPx, color ? `#${color}` : '#595959')
    : null;
  const naturalTextWidth = rich?.width
    ?? Math.max(...lines.map(line => ctx.measureText(line).width));
  const textStyle: DataLabelTextStyle = {
    fontColor: tl.labelFontColor ?? undefined,
    fontItalic: italic,
    fontPaintAuthored: tl.labelFontPaintAuthored ?? undefined,
    fontHidden: tl.labelFontHidden ?? undefined,
    fontLanguage: tl.labelFontLanguage ?? undefined,
    fontBaseline: tl.labelFontBaseline ?? undefined,
    textRotation: tl.labelTextRotation ?? undefined,
    textWrap: tl.labelTextWrap ?? undefined,
    textVerticalAnchor: tl.labelTextVerticalAnchor ?? undefined,
    textVerticalMode: tl.labelTextVerticalMode ?? undefined,
    textLInsEmu: tl.labelTextLInsEmu ?? undefined,
    textTInsEmu: tl.labelTextTInsEmu ?? undefined,
    textRInsEmu: tl.labelTextRInsEmu ?? undefined,
    textBInsEmu: tl.labelTextBInsEmu ?? undefined,
    textBodyAuthored: tl.labelTextBodyAuthored ?? undefined,
  };
  const insets = dataLabelInsets(textStyle, ptToPx);
  const naturalWidth = naturalTextWidth + insets.left + insets.right;
  const naturalHeight = (rich?.height ?? lines.length * lineHeight)
    + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    naturalWidth, naturalHeight,
    tl.labelTextRotation ?? undefined,
    tl.labelTextVerticalMode ?? undefined,
  );
  const placement = placeTrendlineLabel(
    chartRect,
    plotRect,
    rotated.w,
    rotated.h,
    fontPx,
    tl.labelManualLayout,
    labelContext.automaticAnchor,
  );
  if (!placement) return;

  ctx.save();
  if (placement.automatic) {
    ctx.beginPath();
    ctx.rect(plotRect.x, plotRect.y, plotRect.w, plotRect.h);
    ctx.clip();
  }
  const centerX = placement.x + placement.w / 2;
  const centerY = placement.y + placement.h / 2;
  const hasTextBody = tl.labelTextBodyAuthored === true
    || tl.labelTextRotation != null
    || tl.labelTextWrap != null
    || tl.labelTextVerticalAnchor != null
    || tl.labelTextVerticalMode != null
    || tl.labelTextLInsEmu != null
    || tl.labelTextTInsEmu != null
    || tl.labelTextRInsEmu != null
    || tl.labelTextBInsEmu != null;
  // `manualLayout` sizes the authored label shape. Automatic labels are sized
  // from measured text. `bodyPr@rot` rotates text inside that shape, not the
  // shape paint itself (ECMA-376 §20.1.10.83/§21.2.2.216).
  const boxRect = placement.automatic
    ? {
        x: centerX - naturalWidth / 2,
        y: centerY - naturalHeight / 2,
        w: naturalWidth,
        h: naturalHeight,
      }
    : { x: placement.x, y: placement.y, w: placement.w, h: placement.h };
  // chartStyleRoleTrendlineLabel materializes the effective box before paint.
  // Reapplying the linked role here would reinterpret an already-resolved
  // fail-closed paint as a fresh direct noFill and could revive the fallback.
  const labelBox = tl.labelBox;
  paintChartLabelBox(
    ctx,
    labelBox,
    boxRect,
    ptToPx,
    labelContext.shapeRotationDeg ?? 0,
  );
  const alignment = tl.labelTextAlign;
  ctx.textAlign = alignment === 'r' ? 'right' : alignment === 'ctr' ? 'center' : 'left';
  ctx.textBaseline = 'top';
  ctx.fillStyle = color ? `#${color}` : '#595959';
  const maxTextWidth = Math.max(0, boxRect.w - insets.left - insets.right);
  const maxTextHeight = Math.max(0, boxRect.h - insets.top - insets.bottom);
  const automaticNaturalBlockFits = placement.automatic
    && rotated.radians === 0
    && placement.w === rotated.w
    && placement.h === rotated.h
    && (tl.labelTextWrap == null || tl.labelTextWrap === 'none');
  // When the automatic box was measured from these exact generated lines,
  // preserve them directly. Re-fitting an exact two-line body through a
  // floating-point height division could floor 1.999… to one line and drop
  // the authored R² output even though the measured box had room for both.
  const displayLines = rich ? [] : hasTextBody && !automaticNaturalBlockFits
    ? fitStyledDataLabelLines(
        lines.join('\n'), maxTextWidth, maxTextHeight, lineHeight,
        value => ctx.measureText(value).width, textStyle,
      )
    : lines;
  if (!rich && displayLines.length === 0) {
    ctx.restore();
    return;
  }
  const textX = !hasTextBody
    ? (ctx.textAlign === 'right'
      ? placement.x + placement.w
      : ctx.textAlign === 'center' ? placement.x + placement.w / 2 : placement.x)
    : ctx.textAlign === 'right'
    ? boxRect.x + boxRect.w - insets.right
    : ctx.textAlign === 'center'
      ? boxRect.x + (boxRect.w + insets.left - insets.right) / 2
      : boxRect.x + insets.left;
  const baselineShift = (tl.labelFontBaseline ?? 0) * fontPx;
  const textTop = !hasTextBody
    ? placement.y
    : tl.labelTextVerticalAnchor === 'b'
    ? boxRect.y + boxRect.h - insets.bottom - (rich?.height ?? displayLines.length * lineHeight)
    : tl.labelTextVerticalAnchor === 'ctr'
      ? boxRect.y
        + (boxRect.h - (rich?.height ?? displayLines.length * lineHeight)
          + insets.top - insets.bottom) / 2
      : boxRect.y + insets.top;
  const completeLines = placement.automatic
    ? displayLines.length
    : Math.min(displayLines.length, Math.floor(maxTextHeight / lineHeight));
  if (rotated.radians !== 0) {
    ctx.translate(centerX, centerY);
    ctx.rotate(rotated.radians);
    ctx.translate(-centerX, -centerY);
  }
  if (rich) {
    paintRichDataLabelBlock(
      ctx, rich, textX, textTop, ctx.textAlign, 'top',
      Math.max(rich.width, maxTextWidth),
    );
  } else if (!(tl.labelFontPaintAuthored === true
    && (tl.labelFontHidden === true || tl.labelFontColor == null))) {
    for (let index = 0; index < completeLines; index++) {
      ctx.fillText(
        hasTextBody && textStyle.textWrap === 'none'
          ? displayLines[index]
          : elideToWidth(ctx, displayLines[index], Math.max(0, maxTextWidth || naturalTextWidth)),
        textX,
        textTop + index * lineHeight - baselineShift,
      );
    }
  }
  ctx.restore();
}


/** Draw a series' `<c:trendline>` regression lines (ECMA-376 §21.2.2.211).
 *  Each trendline is fitted over the series' non-null `(categoryIndex, value)`
 *  points via {@link fitTrendline} and stroked through the chart's
 *  `toX` (category-index → pixel) and `toY` (value → pixel) maps. `forward` /
 *  `backward` extend the linear fit past the data ends by that many category
 *  units. Nonlinear types are sampled by the same bounded fitter. `seriesColor`
 *  is the fallback stroke when the trendline declares no
 *  `<a:ln>` color. Byte-stable no-op for series with no trendline. */
export function drawSeriesTrendlines(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  seriesColor: string,
  toX: (i: number) => number,
  toY: (v: number) => number,
  ptToPx: number,
  xValues?: readonly (number | null)[],
  labelContext?: TrendlineLabelContext,
  mapPoint?: (categoryValue: number, seriesValue: number) => { x: number; y: number },
): void {
  const tls = s.trendLines;
  if (!tls || tls.length === 0) return;
  // Collect the fittable (index, value) points once.
  const xs: number[] = []; const ys: number[] = [];
  for (let i = 0; i < s.values.length; i++) {
    const v = s.values[i];
    const x = xValues ? xValues[i] : i;
    if (v != null && x != null && Number.isFinite(v) && Number.isFinite(x)) {
      xs.push(x);
      ys.push(v);
    }
  }
  if (xs.length < 2) return;
  const prevDash = ctx.getLineDash ? ctx.getLineDash() : [];
  for (const tl of tls) {
    const fit = fitTrendline(xs, ys, tl.trendlineType, {
      period: tl.period,
      order: tl.order,
      intercept: tl.intercept,
      forward: tl.forward,
      backward: tl.backward,
    });
    if (fit.xs.length < 2) continue;
    if (![...fit.xs, ...fit.ys].every(Number.isFinite)) continue;
    const candidateStats = tl.trendlineType === 'linear'
      ? linearTrendlineStats(xs, ys, tl.intercept)
      : null;
    const stats = candidateStats && [
      candidateStats.slope,
      candidateStats.intercept,
      candidateStats.rSquared,
    ].every(Number.isFinite)
      ? candidateStats
      : null;
    // For a linear fit, forward/backward extend the two endpoints along the
    // fitted slope (in category-index units).
    let fxs = fit.xs; let fys = fit.ys;
    if (tl.trendlineType === 'linear') {
      const m = (fit.ys[1] - fit.ys[0]) / ((fit.xs[1] - fit.xs[0]) || 1);
      const bwd = tl.backward ?? 0; const fwd = tl.forward ?? 0;
      const x0 = fit.xs[0] - bwd; const x1 = fit.xs[1] + fwd;
      fxs = [x0, x1];
      fys = [fit.ys[0] - m * bwd, fit.ys[1] + m * fwd];
    }
    if (![...fxs, ...fys].every(Number.isFinite)) continue;
    const mapped = fxs.map((x, index) => mapPoint
      ? mapPoint(x, fys[index])
      : ({ x: toX(x), y: toY(fys[index]) }));
    if (!mapped.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) continue;
  if (!tl.lineHidden && (tl.linePaintAuthored !== true || tl.lineColor != null)) {
      if (labelContext?.clipLineToPlot) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(
          labelContext.plotRect.x,
          labelContext.plotRect.y,
          labelContext.plotRect.w,
          labelContext.plotRect.h,
        );
        ctx.clip();
      }
      ctx.strokeStyle = tl.lineColor ? `#${tl.lineColor}` : seriesColor;
      ctx.lineWidth = tl.lineWidthEmu ? axisLineWidthPx(tl.lineWidthEmu, ptToPx) : 1.5;
      // DrawingML line presets are authored paint; omission means solid.
      ctx.setLineDash(dashPatternForPreset(tl.lineDash ?? undefined, ctx.lineWidth));
      ctx.beginPath();
      for (let i = 0; i < mapped.length; i++) {
        const { x: px, y: py } = mapped[i];
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.stroke();
      if (labelContext?.clipLineToPlot) ctx.restore();
    }
    drawTrendlineLabel(ctx, tl, stats, ptToPx, labelContext ? {
      ...labelContext,
      automaticAnchor: mapped.at(-1),
    } : undefined, s.valFormatCode);
  }
  ctx.setLineDash(prevDash);
}


/** Office adds every visible trendline to a series-driven legend. The legend
 * entry is a line key whose authored paint comes from `<c:trendline><c:spPr>`;
 * when `<c:name>` is absent, the application-generated label combines the
 * localized trendline kind with the source-series name. We use the invariant
 * OOXML kind names here and keep an authored name verbatim. */
export function trendlineLegendSeries(series: readonly ChartSeries[]): ChartSeries[] {
  const kindLabel = (kind: string): string => {
    switch (kind) {
      case 'exp': return 'Exponential';
      case 'log': return 'Logarithmic';
      case 'poly': return 'Polynomial';
      case 'power': return 'Power';
      case 'movingAvg': return 'Moving Average';
      default: return 'Linear';
    }
  };
  const entries: ChartSeries[] = [];
  for (const source of series) {
    for (const trendline of source.trendLines ?? []) {
      if (trendline.lineHidden === true
        || (trendline.linePaintAuthored === true && trendline.lineColor == null)) continue;
      const fallbackColor = source.lineColor ?? source.color;
      entries.push({
        name: trendline.name
          ?? `${kindLabel(trendline.trendlineType)} (${source.name || 'Series'})`,
        color: trendline.lineColor ?? fallbackColor,
        lineColor: trendline.lineColor ?? fallbackColor,
        lineWidthEmu: trendline.lineWidthEmu,
        lineHidden: false,
        chartexStyle: { lineDash: trendline.lineDash },
        values: [],
        seriesType: 'line',
        showMarker: false,
      });
    }
  }
  return entries;
}


export function legendSeriesWithTrendlines(chart: ChartModel): ChartSeries[] {
  if (legendIsCategoryDriven(
    chart.chartType,
    chart.series.length,
    chart.varyColors !== false,
  ) || chartVariesColorsByPoint(chart)) {
    return chart.series;
  }
  return chart.series.flatMap(series => [series, ...trendlineLegendSeries([series])]);
}
