// Classic chart markers helpers.
import type { ChartModel, ChartSeries } from '../../types/chart';
import type { Fill } from '../../types/common';
import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import { chartDataPointStyleRole, chartSeriesSourceIndex, chartSeriesVariesByPoint, chartStyleDashChoice, rawLinkedChartStyleRole } from '../effective-style.js';
import { chartStyleDirectFillDecision, chartStyleDirectLineDecision, chartStyleFillCascade, chartStyleLineCascade } from '../style-paint.js';
import { axisLineWidthPx } from '../axis-style.js';
import { seriesHasMarkerDetail } from '../marker-style.js';
import { resolveFill } from '../../shape/paint.js';
import { paintChartImageFill } from '../image-fill.js';
import { dashPatternForLine } from './geometry.js';


// Three fixed gradients with 4 + 5 + 6 stops. Keep the work count aligned
// with the complete material so a bubble is admitted or rejected atomically.
export const BUBBLE_3D_MATERIAL_COMPONENTS = 15;


/** ECMA-376 §21.2.2.21 only enables `bubble3D`; it does not define a lighting
 * material. Paint the bounded application-defined material observed in desktop
 * Excel vector output. A single radial envelope cannot independently
 * express the diffuse highlight, right/lower falloff, and narrow lower
 * reflected-light band, so those three components are composited in order.
 * The recipe is normalized to bubble-local coordinates and is therefore shared
 * by every colour, size, and host transform rather than fitted per sample.
 * `source-atop` preserves the authored fill alpha on every pass. */
export function paintBubble3DMaterial(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  sizePx: number,
): void {
  const previousComposite = ctx.globalCompositeOperation;
  const previousFill = ctx.fillStyle;
  ctx.save();
  ctx.clip();
  const paintLayer = (material: CanvasGradient) => {
    ctx.globalCompositeOperation = 'source-atop';
    ctx.fillStyle = material;
    ctx.fillRect(cx - sizePx / 2, cy - sizePx / 2, sizePx, sizePx);
  };

  const diffuseX = cx - sizePx * 0.08;
  const diffuseY = cy - sizePx * 0.17;
  const diffuse = ctx.createRadialGradient(
    diffuseX, diffuseY, 0,
    diffuseX, diffuseY, sizePx * 0.55,
  );
  diffuse.addColorStop(0, 'rgba(255,255,255,0.72)');
  diffuse.addColorStop(0.14, 'rgba(255,255,255,0.48)');
  diffuse.addColorStop(0.38, 'rgba(255,255,255,0.1)');
  diffuse.addColorStop(1, 'rgba(255,255,255,0)');
  paintLayer(diffuse);

  const shadeX = cx - sizePx * 0.08;
  const shadeY = cy - sizePx * 0.18;
  const shade = ctx.createRadialGradient(
    shadeX, shadeY, 0,
    shadeX, shadeY, sizePx * 0.78,
  );
  shade.addColorStop(0, 'rgba(0,0,0,0)');
  shade.addColorStop(0.3, 'rgba(0,0,0,0)');
  shade.addColorStop(0.46, 'rgba(0,0,0,0.22)');
  shade.addColorStop(0.66, 'rgba(0,0,0,0.48)');
  shade.addColorStop(1, 'rgba(0,0,0,0.62)');
  paintLayer(shade);

  // The annulus centre is above-left. Its narrow 0.8--0.95 radius band
  // crosses the lower-left/lower-centre rim while staying clear of the dark
  // lower-right shoulder, matching the material boundary observed in Excel.
  const rimX = cx - sizePx * 0.2;
  const rimY = cy - sizePx * 0.45;
  const lowerRim = ctx.createRadialGradient(
    rimX, rimY, 0,
    rimX, rimY, sizePx,
  );
  lowerRim.addColorStop(0, 'rgba(255,255,255,0)');
  lowerRim.addColorStop(0.76, 'rgba(255,255,255,0)');
  lowerRim.addColorStop(0.82, 'rgba(255,255,255,0.05)');
  lowerRim.addColorStop(0.87, 'rgba(255,255,255,0.12)');
  lowerRim.addColorStop(0.95, 'rgba(255,255,255,0.28)');
  lowerRim.addColorStop(1, 'rgba(255,255,255,0)');
  paintLayer(lowerRim);

  // Recording contexts used by hosts/tests do not necessarily model a full
  // Canvas state stack, so restore the property explicitly as well.
  ctx.globalCompositeOperation = previousComposite;
  ctx.fillStyle = previousFill;
  ctx.restore();
}


export function drawChartMarker(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  cx: number,
  cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
  lineWidthPx: number | undefined,
  fillPaint: Fill | null | undefined,
  shapeRotationDeg: number,
  linePaint: ChartModel['plotAreaLineFill'] | null | undefined = undefined,
  lineDash: string | null | undefined = undefined,
  lineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  lineCap: string | null | undefined = undefined,
  lineJoin: string | null | undefined = undefined,
  bubble3D = false,
  bubble = false,
): void {
  const pointEffect = bubble
    ? chartStyleEffectOwner(point?.chartexStyle)
    // A point marker is the painted CT_DPt shape: its nested marker/spPr is
    // most specific, then dPt/spPr. The series marker/spPr supplies the marker
    // default; series/spPr belongs to the line/area body and must not leak onto
    // every marker.
    : chartStyleEffectOwner(
        point?.markerStyle,
        point?.chartexStyle,
      );
  const seriesEffect = bubble
    ? chartStyleEffectOwner(series.chartexStyle)
    : chartStyleEffectOwner(series.markerStyle);
  const directEffect = pointEffect ?? seriesEffect;
  const sourceSeriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const seriesStyleIndex = series.chartexFormatIdx
    ?? sourceSeriesIndex;
  const variesByPoint = chartSeriesVariesByPoint(chart, sourceSeriesIndex);
  const directEffectIndex = pointEffect ? pointIndex : seriesStyleIndex;
  const linkedMarkerStyle = chartDataPointStyleRole(
    chart, 'dataPointMarker', sourceSeriesIndex,
  );
  const rawLinkedMarkerStyle = rawLinkedChartStyleRole(chart, 'dataPointMarker');
  const linkedStyleIndex = variesByPoint ? pointIndex : seriesStyleIndex;

  const pointFillDecision = chartStyleDirectFillDecision(
    point?.markerStyle, rawLinkedMarkerStyle, pointIndex,
  );
  const seriesFillDecision = chartStyleDirectFillDecision(
    series.markerStyle, rawLinkedMarkerStyle, seriesStyleIndex,
  );
  const pointOwnsFill = pointFillDecision !== undefined
    || point?.markerFill != null
    || point?.markerFillPaintAuthored === true && point.markerStyle?.fillHidden !== true;
  const seriesOwnsFill = seriesFillDecision !== undefined
    || series.markerFill != null
    || series.markerFillPaintAuthored === true && series.markerStyle?.fillHidden !== true;
  let effectiveFill = fill;
  let effectiveFillPaint = fillPaint;
  if (!bubble && !pointOwnsFill && !seriesOwnsFill) {
    const directFillOwner = point?.markerStyle?.shapePropertiesPresent === true
      ? point.markerStyle
      : series.markerStyle;
    const linkedFillDecision = chartStyleFillCascade(
      linkedMarkerStyle,
      rawLinkedMarkerStyle,
      linkedStyleIndex,
      directFillOwner,
    );
    if (linkedFillDecision === null) {
      effectiveFill = '00000000';
      effectiveFillPaint = null;
    } else if (linkedFillDecision?.fillType === 'solid') {
      effectiveFill = linkedFillDecision.color;
      effectiveFillPaint = undefined;
    } else if (linkedFillDecision !== undefined) {
      effectiveFillPaint = linkedFillDecision;
    }
  }
  const pointLineDecision = chartStyleDirectLineDecision(
    point?.markerStyle, rawLinkedMarkerStyle, pointIndex,
  );
  const seriesLineDecision = chartStyleDirectLineDecision(
    series.markerStyle, rawLinkedMarkerStyle, seriesStyleIndex,
  );
  let effectiveLine = line;
  let effectiveLinePaint = linePaint;
  if (!bubble && effectiveLinePaint === undefined) {
    if (pointLineDecision !== undefined) {
      effectiveLinePaint = pointLineDecision?.fillType === 'solid' ? undefined : pointLineDecision;
    }
    else if (point?.markerLine != null) effectiveLinePaint = undefined;
    else if (point?.markerLinePaintAuthored === true
      && point.markerStyle?.lineHidden !== true
      && (point.markerLine == null || point.markerLine === '00000000')) effectiveLinePaint = null;
    else if (seriesLineDecision !== undefined) {
      effectiveLinePaint = seriesLineDecision?.fillType === 'solid' ? undefined : seriesLineDecision;
    }
    else if (series.markerLine != null) effectiveLinePaint = undefined;
    else if (series.markerLinePaintAuthored === true
      && series.markerStyle?.lineHidden !== true
      && (series.markerLine == null || series.markerLine === '00000000')) effectiveLinePaint = null;
    else {
      const directLineOwner = point?.markerStyle?.shapePropertiesPresent === true
        ? point.markerStyle
        : series.markerStyle;
      const linkedLineDecision = chartStyleLineCascade(
        linkedMarkerStyle,
        rawLinkedMarkerStyle,
        linkedStyleIndex,
        directLineOwner,
      );
      if (linkedLineDecision?.fillType === 'solid') {
        effectiveLine = linkedLineDecision.color;
        effectiveLinePaint = undefined;
      } else {
        effectiveLinePaint = linkedLineDecision;
        if (linkedLineDecision === null) effectiveLine = null;
      }
    }
  }
  const linkedLineGeometry = bubble ? undefined : linkedMarkerStyle;
  const effectiveLineWidthPx = lineWidthPx ?? (() => {
    const widthEmu = point?.markerStyle?.lineWidthEmu
      ?? series.markerStyle?.lineWidthEmu ?? linkedLineGeometry?.lineWidthEmu;
    return widthEmu != null ? axisLineWidthPx(widthEmu, ptToPx) : undefined;
  })();
  const markerDashChoice = chartStyleDashChoice(
    lineDash != null || lineCustomDash != null
      ? { lineDash, lineCustomDash, lineDashAuthored: true }
      : undefined,
    point?.markerStyle,
    series.markerStyle,
    linkedLineGeometry,
  );
  const effectiveLineDash = markerDashChoice?.lineDash;
  const effectiveLineCustomDash = markerDashChoice?.lineCustomDash;
  const effectiveLineCap = lineCap ?? point?.markerStyle?.lineCap
    ?? series.markerStyle?.lineCap ?? linkedLineGeometry?.lineCap;
  const effectiveLineJoin = lineJoin ?? point?.markerStyle?.lineJoin
    ?? series.markerStyle?.lineJoin ?? linkedLineGeometry?.lineJoin;
  const fallbackEffect = bubble
    ? chartDataPointStyleRole(
        chart,
        bubble3D ? 'dataPoint3D' : 'dataPoint',
        sourceSeriesIndex,
      )
    : linkedMarkerStyle;
  const fallbackEffectIndex = bubble && chartSeriesVariesByPoint(
    chart, sourceSeriesIndex,
  ) ? pointIndex : (variesByPoint ? pointIndex : seriesStyleIndex);
  drawMarker(
    ctx,
    cx, cy,
    symbol,
    sizePt,
    effectiveFill,
    effectiveLine,
    ptToPx,
    effectiveLineWidthPx,
    effectiveFillPaint,
    shapeRotationDeg,
    effectiveLinePaint,
    effectiveLineDash,
    effectiveLineCustomDash,
    effectiveLineCap,
    effectiveLineJoin,
    bubble3D,
    directEffect,
    fallbackEffect,
    directEffectIndex,
    fallbackEffectIndex,
  );
}


/** Linked/numeric marker styles are part of the effective marker even when the
 * series itself has no `<c:marker><c:spPr>`. Route those automatic glyphs
 * through the shared resolver so fill, line, and atomic effect precedence are
 * identical to explicitly formatted markers. This does not make a marker
 * visible for families (notably area) whose own visibility rule disables it. */
export function seriesHasResolvedMarkerDetail(
  chart: ChartModel,
  series: ChartSeries,
  sourceSeriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series)),
): boolean {
  return seriesHasMarkerDetail(series)
    || chartDataPointStyleRole(chart, 'dataPointMarker', sourceSeriesIndex) != null;
}


/** Draw a single ECMA-376 §21.2.2.32 marker shape centered at `(cx, cy)`.
 *  `sizePt` is the spec's marker side length in points (Excel's default
 *  is 5). `fill` and `line` are hex strings; a leading `#` is tolerated so
 *  callers that route through `chartColor` (which returns `#RRGGBB`)
 *  don't end up double-prefixing into an invalid `##RRGGBB`. `line` may
 *  be null in which case no outline is drawn. `picture` uses the host-warmed
 *  image lookup and fails closed when its authored relationship is unresolved. */
export function drawMarker(
  ctx: CanvasRenderingContext2D,
  cx: number, cy: number,
  symbol: string,
  sizePt: number,
  fill: string,
  line: string | null,
  ptToPx: number,
  lineWidthPx: number = 1,
  /** undefined uses `fill`; null is authored noFill. */
  fillPaint: Fill | null | undefined = undefined,
  shapeRotationDeg = 0,
  /** undefined uses `line`; null is authored line noFill. */
  linePaint: ChartModel['plotAreaLineFill'] | null | undefined = undefined,
  lineDash: string | null | undefined = undefined,
  lineCustomDash: ChartModel['plotAreaLineCustomDash'] = undefined,
  lineCap: string | null | undefined = undefined,
  lineJoin: string | null | undefined = undefined,
  bubble3D = false,
  /** Direct marker/point effect component. */
  effectDirect: import('../../types/chart.js').ChartExElementStyle | null | undefined = undefined,
  /** Resolved linked/numeric `dataPointMarker` or `dataPoint3D` role. */
  effectFallback: import('../../types/chart.js').ChartExElementStyle | null | undefined = undefined,
  effectIndex = 0,
  effectFallbackIndex = effectIndex,
): void {
  const sizePx = Math.max(2, sizePt * ptToPx);
  const half = sizePx / 2;
  if (effectDirect !== undefined || effectFallback !== undefined) {
    paintChartStyleEffects(
      ctx,
      effectDirect,
      effectFallback,
      effectIndex,
      { x: cx - half, y: cy - half, w: sizePx, h: sizePx },
      ptToPx,
      target => drawMarker(
        target,
        cx, cy,
        symbol,
        sizePt,
        fill,
        line,
        ptToPx,
        lineWidthPx,
        fillPaint,
        shapeRotationDeg,
        linePaint,
        lineDash,
        lineCustomDash,
        lineCap,
        lineJoin,
        bubble3D,
      ),
      effectFallbackIndex,
    );
    return;
  }
  const fillCss = fill.startsWith('#') ? fill : `#${fill}`;
  const lineCss = line ? (line.startsWith('#') ? line : `#${line}`) : null;
  ctx.save();
  ctx.fillStyle = fillPaint === undefined
    ? fillCss
    : (fillPaint == null
        ? 'rgba(0,0,0,0)'
        : resolveFill(
            fillPaint, ctx, cx - half, cy - half, sizePx, sizePx, shapeRotationDeg,
          ) ?? 'rgba(0,0,0,0)');
  const resolvedLineStyle = linePaint === undefined
    ? lineCss
    : linePaint == null
      ? null
      : resolveFill(
          linePaint, ctx, cx - half, cy - half, sizePx, sizePx, shapeRotationDeg,
        );
  const hasLine = resolvedLineStyle != null;
  if (resolvedLineStyle) {
    ctx.strokeStyle = resolvedLineStyle;
    ctx.lineWidth = lineWidthPx;
    ctx.setLineDash(dashPatternForLine(lineCustomDash, lineDash, lineWidthPx));
    ctx.lineCap = lineCap === 'rnd' ? 'round' : lineCap === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = lineJoin === 'round' || lineJoin === 'bevel' ? lineJoin : 'miter';
  }
  const imageFill = fillPaint?.fillType === 'image' ? fillPaint : undefined;
  const fillCurrentPath = () => {
    if (!imageFill) {
      if (fillPaint !== null) ctx.fill();
      return;
    }
    ctx.save();
    ctx.clip();
    paintChartImageFill(
      ctx, imageFill, cx - half, cy - half, sizePx, sizePx, ptToPx, shapeRotationDeg,
    );
    ctx.restore();
  };
  const paintMaterial = () => {
    if (bubble3D && fillPaint !== null) paintBubble3DMaterial(ctx, cx, cy, sizePx);
  };
  switch (symbol) {
    case 'square': {
      if (imageFill || bubble3D) {
        ctx.beginPath();
        ctx.rect(cx - half, cy - half, sizePx, sizePx);
        fillCurrentPath();
        paintMaterial();
      } else if (fillPaint !== null) {
        ctx.fillRect(cx - half, cy - half, sizePx, sizePx);
      }
      if (hasLine) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      break;
    }
    case 'diamond': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy);
      ctx.lineTo(cx, cy + half);
      ctx.lineTo(cx - half, cy);
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'triangle': {
      ctx.beginPath();
      ctx.moveTo(cx, cy - half);
      ctx.lineTo(cx + half, cy + half);
      ctx.lineTo(cx - half, cy + half);
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'x': {
      ctx.strokeStyle = resolvedLineStyle ?? ctx.fillStyle;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy - half); ctx.lineTo(cx + half, cy + half);
      ctx.moveTo(cx - half, cy + half); ctx.lineTo(cx + half, cy - half);
      ctx.stroke();
      break;
    }
    case 'plus': {
      ctx.strokeStyle = resolvedLineStyle ?? ctx.fillStyle;
      ctx.lineWidth = Math.max(1, sizePx * 0.18);
      ctx.beginPath();
      ctx.moveTo(cx - half, cy); ctx.lineTo(cx + half, cy);
      ctx.moveTo(cx, cy - half); ctx.lineTo(cx, cy + half);
      ctx.stroke();
      break;
    }
    case 'star': {
      // 5-point star inscribed in a circle of radius `half`.
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? half : half * 0.45;
        const a = -Math.PI / 2 + i * Math.PI / 5;
        const px = cx + Math.cos(a) * r;
        const py = cy + Math.sin(a) * r;
        if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      ctx.closePath();
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'dot': {
      // ECMA-376 §21.2.3.27: width=1/2 and height=1/5 of marker size.
      ctx.beginPath();
      ctx.ellipse(cx, cy, sizePx * 0.25, sizePx * 0.1, 0, 0, Math.PI * 2);
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
    case 'dash': {
      // ECMA-376 §21.2.3.27: height=1/5 of marker size.
      const dh = sizePx * 0.2;
      if (imageFill || bubble3D) {
        ctx.beginPath(); ctx.rect(cx - half, cy - dh / 2, sizePx, dh); fillCurrentPath();
        paintMaterial();
      } else if (fillPaint !== null) {
        ctx.fillRect(cx - half, cy - dh / 2, sizePx, dh);
      }
      if (hasLine) ctx.strokeRect(cx - half, cy - dh / 2, sizePx, dh);
      break;
    }
    case 'picture': {
      ctx.beginPath();
      ctx.rect(cx - half, cy - half, sizePx, sizePx);
      if (imageFill) {
        paintChartImageFill(
          ctx, imageFill, cx - half, cy - half, sizePx, sizePx, ptToPx, shapeRotationDeg,
        );
      }
      paintMaterial();
      // Fill and line are independent CT_ShapeProperties components. An
      // authored noFill/unresolved blip must not suppress the picture outline.
      if (hasLine) ctx.strokeRect(cx - half, cy - half, sizePx, sizePx);
      ctx.restore();
      return;
    }
    case 'circle':
    default: {
      ctx.beginPath();
      ctx.arc(cx, cy, half, 0, Math.PI * 2);
      fillCurrentPath();
      paintMaterial();
      if (hasLine) ctx.stroke();
      break;
    }
  }
  ctx.restore();
}
