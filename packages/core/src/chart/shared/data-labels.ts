// Classic chart data labels helpers.
import type { ChartDataLabelOverride, ChartDisplayUnits, ChartLabelBox, ChartModel, ChartRect, ChartSeries, ChartTextRun } from '../../types/chart';
import { boundDataLabelText, resolveDataLabelPlacement } from '../data-label-layout.js';
import type { DataLabelAnchor, DataLabelRect } from '../data-label-layout.js';
import { paintRichDataLabelBlock, resolveRichDataLabelBlock } from '../rich-data-label.js';
import type { RichDataLabelOptions } from '../rich-data-label.js';
import { anchoredDataLabelPoint, dataLabelCanvasTextAlign, dataLabelInsets, dataLabelIsDeleted, effectiveDataLabelTextStyle, fitStyledDataLabelLines, rotatedDataLabelSize, transformDataLabelText } from '../data-label-style.js';
import type { DataLabelTextStyle } from '../data-label-style.js';
import { effectiveDataLabelText } from '../data-label-content.js';
import { formatCategoryLabel, formatChartValWithCode } from '../chart-number-format.js';
import { chartTextFontSizePx } from '../layout.js';
import { mergeChartLabelBoxes, paintChartLabelBox } from '../label-box.js';
import { chartImageFillPaintWorkUpperBound } from '../image-fill.js';
import type { ChartImageLookup } from '../image-fill.js';
import { classicDataLabelPointIsPainted, markerPaintComponents } from '../marker-style.js';
import type { ChartThreeDRenderer } from '../three-d-contract.js';
import { PT_TO_PX } from '../../units.js';
import { LEGEND_SWATCH_TEXT_GAP, drawLegendSwatch, legendSwatchHeight, legendSwatchWidths } from './legend.js';
import type { DataLabelLegendKey } from './legend.js';
import { indexPointOverrides } from './palette.js';
import { scatterXValue } from './scatter-geometry.js';
import { displayUnitDivisor } from './axis.js';
import { chartFontFamily } from './fonts.js';
import { CLASSIC_THREE_D_FAMILIES, MAX_CANVAS_LABEL_GRADIENT_STOPS, MAX_CANVAS_LABEL_PAINT_COMPONENTS } from './paint-limits.js';


/** ECMA-376 §21.2.2.180: omission/false suppresses a label whose data-point value
 * is numerically greater than the effective value-axis maximum. This gate runs
 * after the shared axis planner has resolved authored and automatic bounds; it
 * never changes those bounds. For stacked charts the comparison remains the
 * point's authored value; the cumulative stack endpoint is layout geometry,
 * not the value represented by that label. */
export function dataLabelWithinAxisMaximum(
  chart: Pick<ChartModel, 'showDataLabelsOverMax'>,
  plottedValue: number,
  axisMaximum: number,
): boolean {
  return chart.showDataLabelsOverMax === true
    || !Number.isFinite(axisMaximum)
    || plottedValue <= axisMaximum;
}


/** Draw per-point data labels: position-aware text near each marker. */
export function drawSeriesDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  useIndexX: boolean,
  toX: (v: number) => number,
  toY: (v: number) => number,
  ph: number,
  ptToPx: number,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). Threaded so date-format
   *  value labels resolve against the correct epoch. Defaults to false, which
   *  also accepts the optional `ChartModel.date1904` when it is undefined. */
  date1904 = false,
  /** Resolved data-label CSS font-family; defaults to sans-serif (byte-stable). */
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default (scatter defaults to `'r'`). */
  defaultPos = 'r',
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  layoutReferenceRect: DataLabelRect = bounds,
  richFontFamilyForFace?: (face: string) => string,
  valueDisplayUnits?: ChartDisplayUnits | null,
  legendKeyAt?: (pointIndex: number) => DataLabelLegendKey | undefined,
  isValueVisible?: (value: number) => boolean,
  shapeRotationDeg = 0,
  markerGapAt?: (pointIndex: number) => number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  if (overrides.length === 0 && !s.seriesDataLabels) return;
  const seriesDef = s.seriesDataLabels;
  for (let i = 0; i < s.values.length; i++) {
    const yv = s.values[i]; if (yv == null) continue;
    if (isValueVisible && !isValueVisible(yv)) continue;
    const xv = scatterXValue(cats, i, useIndexX);
    if (xv == null) continue;
    const ovr = overridesByIndex.get(i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the point; a per-point
    // `<c:dLbl>` that only carries style / flag overrides (empty `<c:tx>`) is NOT
    // a delete — key off the explicit `deleted` flag, then honor per-point
    // show-flags (§21.2.2.47) over the series defaults.
    if (dataLabelIsDeleted(seriesDef, ovr)) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    const showBubbleSize = ovr?.showBubbleSize ?? seriesDef?.showBubbleSize;
    const showLegendKey = ovr?.showLegendKey ?? seriesDef?.showLegendKey ?? false;
    const text = effectiveDataLabelText({
      customText: ovr?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showBubbleSize,
      category: useIndexX
        ? formatCategoryLabel(
          (cats[i] ?? String(xv)).toString(),
          s.catFormatCodes?.[i] ?? s.catFormatCode ?? null,
          date1904,
        )
        : formatChartValWithCode(
          xv, s.catFormatCodes?.[i] ?? s.catFormatCode ?? null, date1904,
        ),
      seriesName: s.name,
      sourceValue: yv,
      bubbleSize: s.bubbleSizes?.[i] ?? undefined,
      valueDivisor: displayUnitDivisor(valueDisplayUnits),
      formatCode: ovr?.formatCode ?? seriesDef?.formatCode ?? null,
      date1904,
      separator: ovr?.separator ?? seriesDef?.separator,
    });
    const legendKey = showLegendKey ? legendKeyAt?.(i) : undefined;
    if (!text && !legendKey) continue;
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = chartTextFontSizePx(sizeHpt, ptToPx)
      ?? Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    const labelFace = ovr?.fontFace ?? seriesDef?.fontFace;
    const labelFont = labelFace && richFontFamilyForFace
      ? richFontFamilyForFace(labelFace)
      : fontFamily;
    drawDataLabelText(
      ctx, toX(xv), toY(yv), text, pos, fontSizePx, color, bold, labelFont,
      markerGapAt?.(i) ?? 0,
      bounds, ovr?.manualLayout,
      layoutReferenceRect,
      ovr?.richRuns,
      ptToPx,
      richFontFamilyForFace,
      legendKey,
      effectiveDataLabelTextStyle(ovr, seriesDef),
      mergeChartLabelBoxes(ovr?.labelBox, seriesDef?.labelBox),
      shapeRotationDeg,
    );
  }
}


export function drawDataLabelText(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  text: string,
  position: string,
  fontSizePx: number,
  color: string | undefined,
  bold: boolean,
  fontFamily = 'sans-serif',
  /** Extra gap (px) added to the text offset in the label's direction so the
   *  text clears an anchor glyph (e.g. a line-chart marker). The shared base
   *  inset is one half-em; markerGap is added outside that inset. */
  markerGap = 0,
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  manualLayout?: ChartDataLabelOverride['manualLayout'],
  layoutReferenceRect: DataLabelRect = bounds,
  richRuns?: readonly ChartTextRun[],
  ptToPx = 1,
  richFontFamilyForFace?: (face: string) => string,
  legendKey?: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  labelBox?: ChartLabelBox,
  shapeRotationDeg = 0,
): void {
  ctx.save();
  ctx.font = `${textStyle?.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
  drawBoundedDataLabelText(
    ctx,
    text,
    { kind: 'point', x: cx, y: cy, position, markerGap },
    bounds,
    fontSizePx,
    color ? `#${color}` : '#333',
    manualLayout,
    layoutReferenceRect,
    richRuns && richRuns.length > 0
      ? {
          runs: richRuns,
          ptToPx,
          fontFamily,
          fallbackBold: bold,
          fallbackItalic: textStyle?.fontItalic,
          fallbackBaseline: textStyle?.fontBaseline,
          fallbackColorHidden: textStyle?.fontPaintAuthored === true
            && (textStyle.fontHidden === true || textStyle.fontColor == null),
          fontFamilyForFace: richFontFamilyForFace,
        }
      : undefined,
    legendKey,
    textStyle,
    ptToPx,
    labelBox,
    shapeRotationDeg,
  );
  ctx.restore();
}


/** A `<c:tx><c:rich>` body is authoritative only with non-empty custom text.
 * Empty override text means the visible label is composed from show/format
 * flags, so stale/empty rich payload must not replace that composition. */
export function customRichDataLabelOptions(
  chart: ChartModel,
  override: ChartDataLabelOverride | undefined,
  ptToPx: number,
  fontFamily: string,
  fallbackBold: boolean,
  textStyle?: DataLabelTextStyle,
): RichDataLabelOptions | undefined {
  if (!override?.text || !override.richRuns || override.richRuns.length === 0) return undefined;
  return richDataLabelOptions(
    chart, override.richRuns, ptToPx, fontFamily, fallbackBold, textStyle,
  );
}


export function richDataLabelOptions(
  chart: ChartModel,
  runs: ChartDataLabelOverride['richRuns'],
  ptToPx: number,
  fontFamily: string,
  fallbackBold: boolean,
  textStyle?: DataLabelTextStyle,
): RichDataLabelOptions | undefined {
  if (!runs || runs.length === 0) return undefined;
  return {
    runs,
    ptToPx,
    fontFamily,
    fallbackBold,
    fallbackItalic: textStyle?.fontItalic,
    fallbackBaseline: textStyle?.fontBaseline,
    fallbackColorHidden: textStyle?.fontPaintAuthored === true
      && (textStyle.fontHidden === true || textStyle.fontColor == null),
    fontFamilyForFace: face => chartFontFamily(chart, face, 'minor'),
  };
}


/** Measure, fit, clip, and paint one label through the shared pure resolver. */
export function drawBoundedDataLabelText(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchor: DataLabelAnchor,
  bounds: DataLabelRect,
  fontSizePx: number,
  color: string,
  manualLayout?: ChartDataLabelOverride['manualLayout'],
  layoutReferenceRect: DataLabelRect = bounds,
  rich?: RichDataLabelOptions,
  legendKey?: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  textPtToPx = 1,
  labelBox?: ChartLabelBox,
  shapeRotationDeg = 0,
): void {
  if ((!text && !legendKey) || !Number.isFinite(fontSizePx) || fontSizePx <= 0) return;
  if (legendKey) {
    drawBoundedDataLabelWithLegendKey(
      ctx, text, anchor, bounds, fontSizePx, color, manualLayout,
      layoutReferenceRect, rich, legendKey,
      textStyle,
      labelBox,
    );
    return;
  }
  if (rich) {
    const block = resolveRichDataLabelBlock(ctx, rich, fontSizePx, color);
    if (!block) return;
    const insets = dataLabelInsets(textStyle, textPtToPx);
    const rotated = rotatedDataLabelSize(
      block.width + insets.left + insets.right,
      block.height + insets.top + insets.bottom,
      textStyle?.textRotation,
      textStyle?.textVerticalMode,
    );
    const placement = resolveDataLabelPlacement(
      anchor, bounds, { w: rotated.w, h: rotated.h }, fontSizePx, manualLayout,
      layoutReferenceRect,
    );
    if (!placement) return;

    ctx.save();
    ctx.beginPath();
    ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
    ctx.clip();
    paintChartLabelBox(ctx, labelBox, placement.rect, textPtToPx, shapeRotationDeg);
    const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
    const anchored = anchoredDataLabelPoint(
      placement.x, placement.y, placement.rect,
      block.height + insets.top + insets.bottom, textStyle, manualLayout != null,
      paintAlign, placement.textAlign,
      block.width + insets.left + insets.right, rotated.radians,
    );
    const transformed = transformDataLabelText(
      ctx, anchored.x, anchored.y, rotated.radians, paintAlign,
      placement.textBaseline, insets,
    );
    paintRichDataLabelBlock(
      ctx, block, transformed.x, transformed.y, paintAlign, placement.textBaseline,
      manualLayout ? Math.max(0, placement.rect.w - insets.left - insets.right) : block.width,
    );
    ctx.restore();
    return;
  }
  const lineHeight = fontSizePx * 1.15;
  const sourceLines = boundDataLabelText(text).value.split(/\r?\n/);
  const measuredW = sourceLines.reduce((max, line) => Math.max(max, ctx.measureText(line).width), 0);
  const measuredH = Math.max(lineHeight, sourceLines.length * lineHeight);
  const insets = dataLabelInsets(textStyle, textPtToPx);
  const measuredRotated = rotatedDataLabelSize(
    measuredW + insets.left + insets.right,
    measuredH + insets.top + insets.bottom,
    textStyle?.textRotation,
    textStyle?.textVerticalMode,
  );
  let placement = resolveDataLabelPlacement(
    anchor, bounds, { w: measuredRotated.w, h: measuredRotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;
  const measure = (value: string): number => ctx.measureText(value).width;
  const lines = fitStyledDataLabelLines(
    text, placement.maxWidth, placement.maxHeight, lineHeight, measure, textStyle,
  );
  if (lines.length === 0) return;
  const fittedW = lines.reduce((max, line) => Math.max(max, measure(line)), 0);
  const fittedH = lines.length * lineHeight;
  const fittedRotated = rotatedDataLabelSize(
    fittedW + insets.left + insets.right,
    fittedH + insets.top + insets.bottom,
    textStyle?.textRotation,
    textStyle?.textVerticalMode,
  );
  placement = resolveDataLabelPlacement(
    anchor, bounds, { w: fittedRotated.w, h: fittedRotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  paintChartLabelBox(ctx, labelBox, placement.rect, textPtToPx, shapeRotationDeg);
  const textPaintUnavailable = textStyle?.fontPaintAuthored === true
    && (textStyle.fontHidden === true || textStyle.fontColor == null);
  ctx.fillStyle = color;
  const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
  ctx.textAlign = paintAlign;
  ctx.textBaseline = placement.textBaseline;
  const anchored = anchoredDataLabelPoint(
    placement.x, placement.y, placement.rect,
    fittedH + insets.top + insets.bottom, textStyle, manualLayout != null,
    paintAlign, placement.textAlign,
    fittedW + insets.left + insets.right, fittedRotated.radians,
  );
  const transformed = transformDataLabelText(
    ctx, anchored.x, anchored.y, fittedRotated.radians, paintAlign,
    placement.textBaseline, insets,
  );
  const baselineShift = (textStyle?.fontBaseline ?? 0) * fontSizePx;
  const firstY = placement.textBaseline === 'middle'
    ? transformed.y - ((lines.length - 1) * lineHeight) / 2
    : placement.textBaseline === 'bottom'
      ? transformed.y - ((lines.length - 1) * lineHeight)
      : transformed.y;
  if (!textPaintUnavailable) for (let index = 0; index < lines.length; index++) {
    ctx.fillText(lines[index], transformed.x, firstY + index * lineHeight - baselineShift);
  }
  ctx.restore();
}


/** Measure and paint a data-label legend key and its optional text as one
 * bounded block. Existing legend swatch geometry is reused verbatim, while the
 * shared data-label placement resolver owns clipping and manual layout. */
export function drawBoundedDataLabelWithLegendKey(
  ctx: CanvasRenderingContext2D,
  text: string,
  anchor: DataLabelAnchor,
  bounds: DataLabelRect,
  fontSizePx: number,
  color: string,
  manualLayout: ChartDataLabelOverride['manualLayout'] | undefined,
  layoutReferenceRect: DataLabelRect,
  rich: RichDataLabelOptions | undefined,
  legendKey: DataLabelLegendKey,
  textStyle?: DataLabelTextStyle,
  labelBox?: ChartLabelBox,
): void {
  const { entry, ptToPx, shapeRotationDeg } = legendKey;
  const keyWidth = legendSwatchWidths([entry], fontSizePx, ptToPx)[0] ?? 0;
  const keyHeight = legendSwatchHeight(entry, fontSizePx, ptToPx);
  const gap = text ? LEGEND_SWATCH_TEXT_GAP : 0;
  const richBlock = text && rich
    ? resolveRichDataLabelBlock(ctx, rich, fontSizePx, color)
    : null;
  if (text && rich && !richBlock) return;
  const lineHeight = fontSizePx * 1.15;
  const sourceLines = text && !richBlock
    ? boundDataLabelText(text).value.split(/\r?\n/)
    : [];
  const sourceTextWidth = richBlock?.width ?? sourceLines.reduce(
    (max, line) => Math.max(max, ctx.measureText(line).width), 0,
  );
  const sourceTextHeight = richBlock?.height
    ?? (sourceLines.length > 0 ? Math.max(lineHeight, sourceLines.length * lineHeight) : 0);
  const insets = dataLabelInsets(textStyle, ptToPx);
  const sourceWidth = keyWidth + gap + sourceTextWidth + insets.left + insets.right;
  const sourceHeight = Math.max(keyHeight, sourceTextHeight) + insets.top + insets.bottom;
  const sourceRotated = rotatedDataLabelSize(
    sourceWidth, sourceHeight, textStyle?.textRotation, textStyle?.textVerticalMode,
  );
  let placement = resolveDataLabelPlacement(
    anchor,
    bounds,
    { w: sourceRotated.w, h: sourceRotated.h },
    fontSizePx,
    manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  let lines = sourceLines;
  if (text && !richBlock) {
    lines = fitStyledDataLabelLines(
      text,
      Math.max(0, placement.maxWidth - keyWidth - gap),
      placement.maxHeight,
      lineHeight,
      value => ctx.measureText(value).width,
      textStyle,
    );
    if (lines.length === 0) return;
  }
  const textWidth = richBlock?.width ?? lines.reduce(
    (max, line) => Math.max(max, ctx.measureText(line).width), 0,
  );
  const textHeight = richBlock?.height ?? (lines.length * lineHeight);
  const contentWidth = keyWidth + gap + textWidth;
  const contentHeight = Math.max(keyHeight, textHeight);
  const totalWidth = contentWidth + insets.left + insets.right;
  const totalHeight = contentHeight + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    totalWidth, totalHeight, textStyle?.textRotation, textStyle?.textVerticalMode,
  );
  placement = resolveDataLabelPlacement(
    anchor, bounds, { w: rotated.w, h: rotated.h }, fontSizePx, manualLayout,
    layoutReferenceRect,
  );
  if (!placement) return;

  let centerX = placement.textAlign === 'left'
    ? placement.x + rotated.w / 2
    : placement.textAlign === 'right'
      ? placement.x - rotated.w / 2
      : placement.x;
  let centerY = placement.textBaseline === 'top'
    ? placement.y + rotated.h / 2
    : placement.textBaseline === 'bottom'
      ? placement.y - rotated.h / 2
      : placement.y;
  if (manualLayout) {
    const paintAlign = dataLabelCanvasTextAlign(textStyle, 'center');
    const anchored = anchoredDataLabelPoint(
      centerX, centerY, placement.rect, totalHeight, textStyle, true, paintAlign,
    );
    centerX = paintAlign === 'left' ? anchored.x + totalWidth / 2
      : paintAlign === 'right' ? anchored.x - totalWidth / 2 : anchored.x;
    centerY = anchored.y;
  }
  const left = centerX - totalWidth / 2 + insets.left;
  const top = centerY - totalHeight / 2 + insets.top;
  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  paintChartLabelBox(ctx, labelBox, placement.rect, ptToPx, shapeRotationDeg);
  if (rotated.radians !== 0) {
    ctx.translate(centerX, centerY);
    ctx.rotate(rotated.radians);
    ctx.translate(-centerX, -centerY);
  }
  drawLegendSwatch(
    ctx,
    entry.swatchStyle,
    entry.color,
    left,
    top + (contentHeight - keyHeight) / 2,
    keyWidth,
    keyHeight,
    entry.marker,
    entry.fillPaint,
    entry.outlinePaint,
    entry.outlineColor,
    entry.outlineWidthEmu,
    entry.outlineDash,
    entry.outlineCustomDash,
    entry.outlineCap,
    entry.outlineJoin,
    ptToPx,
    shapeRotationDeg,
    entry.directEffect,
    entry.fallbackEffect,
    entry.directEffectIndex,
    entry.fallbackEffectIndex,
  );
  if (text) {
    const textX = left + keyWidth + gap;
    if (richBlock) {
      paintRichDataLabelBlock(
        ctx, richBlock, textX, top + (contentHeight - textHeight) / 2, 'left', 'top',
      );
    } else if (!(textStyle?.fontPaintAuthored === true
      && (textStyle.fontHidden === true || textStyle.fontColor == null))) {
      ctx.fillStyle = color;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'top';
      const baselineShift = (textStyle?.fontBaseline ?? 0) * fontSizePx;
      const firstY = top + (contentHeight - textHeight) / 2 - baselineShift;
      for (let index = 0; index < lines.length; index++) {
        ctx.fillText(lines[index], textX, firstY + index * lineHeight);
      }
    }
  }
  ctx.restore();
}


/** Per-point data labels for a category-axis series (line / area). Consumes the
 *  same `<c:dLbl idx>` overrides and series-level `<c:dLbls>` block scatter does
 *  ({@link drawSeriesDataLabels}), but maps points by CATEGORY INDEX with the
 *  series' plotted value → px mapping. Returns true when it handled the labels
 *  for this series (so the caller skips the family's legacy `showDataLabels`
 *  path), false when the series has no override/series-level label config.
 *
 *  `plotNullAsZero` mirrors the marker loop's dispBlanksAs gate (§21.2.2.42):
 *  a null cell normally has no label (gap/span leave the point unplotted), but
 *  in "zero" mode the blank IS a plotted point (value 0) and gets a label like
 *  any other — the line-chart caller passes `dispBlanks === 'zero'`. The area
 *  caller passes `true` unconditionally: area's fill has always read a blank
 *  cell as 0 (`?? 0`, dispBlanksAs is a no-op for the filled region), so its
 *  per-point labels have likewise always covered every category index. */
export function drawCategoryDataLabels(
  ctx: CanvasRenderingContext2D,
  s: ChartSeries,
  cats: string[],
  n: number,
  xAt: (ci: number) => number,
  yAt: (v: number) => number,
  plotted: (ci: number) => number,
  ph: number,
  ptToPx: number,
  date1904: boolean,
  plotNullAsZero: boolean,
  // Resolved data-label CSS font-family (element face ?? theme body ??
  // sans-serif). Defaults to sans-serif so callers that don't pass it stay
  // byte-stable.
  fontFamily = 'sans-serif',
  /** Fallback `<c:dLblPos>` (§21.2.2.48) when neither the per-point override nor
   *  the series-level block sets one: the chart-level position, else the
   *  per-chart-type default. Line defaults to `'r'` (PowerPoint), area to
   *  `'ctr'`. */
  defaultPos = 't',
  bounds: DataLabelRect = { x: -1e6, y: -1e6, w: 2e6, h: 2e6 },
  layoutReferenceRect: DataLabelRect = bounds,
  percentRatioAt?: (index: number) => number,
  markerGapAt?: (index: number) => number,
  richFontFamilyForFace?: (face: string) => string,
  valueDisplayUnits?: ChartDisplayUnits | null,
  legendKeyAt?: (pointIndex: number) => DataLabelLegendKey | undefined,
  isValueVisible?: (value: number) => boolean,
  shapeRotationDeg = 0,
): boolean {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  const seriesDef = s.seriesDataLabels;
  if (overrides.length === 0 && !seriesDef) return false;
  for (let ci = 0; ci < n; ci++) {
    if (s.sourceHidden?.[ci] === true) continue;
    if (s.values[ci] == null && !plotNullAsZero) continue;
    const anchorValue = plotted(ci);
    if (isValueVisible && !isValueVisible(anchorValue)) continue;
    const sourceValue = s.values[ci] ?? 0;
    const ovr = overridesByIndex.get(ci);
    // Genuine `<c:delete val="1"/>` (§21.2.2.43) skips; a style/flag-only
    // override is not a delete. Per-point show-flags (§21.2.2.47) win over the
    // series defaults.
    if (dataLabelIsDeleted(seriesDef, ovr)) continue;
    const showCatName = ovr?.showCatName ?? seriesDef?.showCatName;
    const showSerName = ovr?.showSerName ?? seriesDef?.showSerName;
    const showVal     = ovr?.showVal ?? seriesDef?.showVal;
    const showPercent = ovr?.showPercent ?? seriesDef?.showPercent;
    const showLegendKey = ovr?.showLegendKey ?? seriesDef?.showLegendKey ?? false;
    const text = effectiveDataLabelText({
      customText: ovr?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: cats[ci] ?? '',
      seriesName: s.name,
      sourceValue,
      valueDivisor: displayUnitDivisor(valueDisplayUnits),
      percentRatio: percentRatioAt?.(ci),
      formatCode: ovr?.formatCode ?? seriesDef?.formatCode ?? null,
      date1904,
      separator: ovr?.separator ?? seriesDef?.separator,
    });
    const legendKey = showLegendKey ? legendKeyAt?.(ci) : undefined;
    if (!text && !legendKey) continue;
    const pos = ovr?.position ?? seriesDef?.position ?? defaultPos;
    const sizeHpt = ovr?.fontSizeHpt ?? seriesDef?.fontSizeHpt;
    const fontSizePx = chartTextFontSizePx(sizeHpt, ptToPx)
      ?? Math.max(9, Math.min(11, ph / 25));
    const color = ovr?.fontColor ?? seriesDef?.fontColor;
    const bold = ovr?.fontBold ?? seriesDef?.fontBold ?? false;
    const labelFace = ovr?.fontFace ?? seriesDef?.fontFace;
    const labelFont = labelFace && richFontFamilyForFace
      ? richFontFamilyForFace(labelFace)
      : fontFamily;
    drawDataLabelText(
      ctx, xAt(ci), yAt(anchorValue), text, pos, fontSizePx, color, bold, labelFont,
      markerGapAt?.(ci) ?? 0,
      bounds, ovr?.manualLayout,
      layoutReferenceRect,
      ovr?.richRuns,
      ptToPx,
      richFontFamilyForFace,
      legendKey,
      effectiveDataLabelTextStyle(ovr, seriesDef),
      mergeChartLabelBoxes(ovr?.labelBox, seriesDef?.labelBox),
      shapeRotationDeg,
    );
  }
  return true;
}


export function chartLabelBoxPaintComponents(
  box: ChartLabelBox | null | undefined,
  imageLookup: ChartImageLookup | undefined,
  destinationWidth: number,
  destinationHeight: number,
  ptToPx: number,
): number | null {
  let total = 0;
  for (const paint of [box?.fillPaint, box?.borderFill]) {
    if (!paint) continue;
    const components = paint.fillType === 'image'
      ? chartImageFillPaintWorkUpperBound(
          paint, imageLookup, destinationWidth, destinationHeight, ptToPx,
        )
      : markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CANVAS_LABEL_GRADIENT_STOPS) {
      return null;
    }
    total += components;
  }
  return total;
}


export function dataLabelHasContent(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
  override: ChartDataLabelOverride | undefined,
): boolean {
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return false;
  return Boolean(
    override?.text
    || (override?.showVal ?? defaults?.showVal ?? chart.showDataLabels)
    || (override?.showCatName ?? defaults?.showCatName)
    || (override?.showSerName ?? defaults?.showSerName)
    || (override?.showPercent ?? defaults?.showPercent)
    || (override?.showBubbleSize ?? defaults?.showBubbleSize)
    || (override?.showLegendKey ?? defaults?.showLegendKey)
  ) && index < Math.max(series.values.length, series.categories?.length ?? 0, chart.categories.length);
}


/** Bound structured label-shape work before any family starts painting. The
 * count follows the shared 2-D/ChartEx label placement and the optional 3-D
 * label path, plus one generated box per visible 2-D trendline label. */
/** @internal Exported for resource-boundary regression tests. */
export function chartLabelPaintWorkCount(
  chart: ChartModel,
  threeD: ChartThreeDRenderer | undefined,
  imageLookup?: ChartImageLookup,
  ptToPx = PT_TO_PX,
  chartRect?: ChartRect,
): number | null {
  // ChartEx hierarchy labels are expanded only by the optional ChartEx
  // renderer, which owns their resource preflight with the hierarchy model.
  if (chart.chartexSunburst || chart.chartexTreemap) return null;
  let total = 0;
  // Label fitting is family- and text-dependent. The chart rectangle is a
  // monotonic upper bound on every generated label box and therefore on tiled
  // drawImage repetitions; a missing rectangle keeps standalone tests and
  // non-rendering callers on a small bounded estimate.
  const destinationWidth = Math.max(1, chartRect?.w ?? 32 * ptToPx);
  const destinationHeight = Math.max(1, chartRect?.h ?? 16 * ptToPx);
  const charge = (box: ChartLabelBox | null | undefined): boolean => {
    const components = chartLabelBoxPaintComponents(
      box, imageLookup, destinationWidth, destinationHeight, ptToPx,
    );
    if (components == null || components > MAX_CANVAS_LABEL_PAINT_COMPONENTS - total) {
      return false;
    }
    total += components;
    return true;
  };
  const threeDLabels = chart.threeD != null && threeD != null
    && CLASSIC_THREE_D_FAMILIES.has(chart.chartType);
  const scatterHasNumericX = chart.series.some(series => {
    const family = series.seriesType ?? chart.chartType;
    return family === 'scatter' && (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  for (let sourceSeriesIndex = 0; sourceSeriesIndex < chart.series.length; sourceSeriesIndex++) {
    const series = chart.series[sourceSeriesIndex]!;
    const overrides = indexPointOverrides(series.dataLabelOverrides);
    const family = series.seriesType ?? chart.chartType;
    const pointCount = Math.max(
      series.values.length, series.categories?.length ?? 0, chart.categories.length,
    );
    for (let index = 0; index < pointCount; index++) {
      const value = series.values[index];
      if (threeDLabels) {
        if (value == null || !Number.isFinite(value)) continue;
        if (chart.showDataLabelsOverMax !== true) {
          const maximum = series.useSecondaryAxis
            ? chart.secondaryValAxis?.max : chart.valMax;
          if (maximum != null && Number.isFinite(maximum) && value > maximum) continue;
        }
      } else if (!classicDataLabelPointIsPainted(
        chart, series, family, index, scatterHasNumericX, sourceSeriesIndex,
      )) {
        continue;
      }
      const override = overrides.get(index);
      if (!dataLabelHasContent(chart, series, index, override)) continue;
      const box = mergeChartLabelBoxes(override?.labelBox, series.seriesDataLabels?.labelBox);
      if (box && !charge(box)) return MAX_CANVAS_LABEL_PAINT_COMPONENTS + 1;
    }
    if (!threeDLabels) for (const trendline of series.trendLines ?? []) {
      const hasLabelContent = trendline.dispEq === true || trendline.dispRSqr === true
        || Boolean(trendline.labelText)
        || trendline.labelRichRuns?.some(run => run.text.length > 0) === true;
      if (hasLabelContent && trendline.labelBox
        && !charge(trendline.labelBox)) return MAX_CANVAS_LABEL_PAINT_COMPONENTS + 1;
    }
  }

  return total;
}
