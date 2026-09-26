// Classic pie chart family.
import type {
  ChartDataLabelOverride,
  ChartLabelBox,
  ChartModel,
  ChartRect,
  ChartSeries,
  ChartSeriesDataLabels,
} from '../../types/chart';

import { classicDataPointFillDecision } from '../classic-data-point-style.js';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import { mergeChartLabelBoxes, paintChartLabelBox } from '../label-box.js';

import {
  dataLabelCanvasTextAlign,
  dataLabelIsDeleted,
  dataLabelInsets,
  effectiveDataLabelTextStyle,
  fitStyledDataLabelLines,
  rotatedDataLabelSize,
  type DataLabelTextStyle,
} from '../data-label-style.js';

import { computeChartFrame, chartTextFontSizePx } from '../layout.js';

import { planOfPieSecondaryIndices } from '../of-pie.js';

import {
  paintRichDataLabelBlock,
  resolveRichDataLabelBlock,
  type RichDataLabelBlock,
} from '../rich-data-label.js';
import { effectiveDataLabelText } from '../data-label-content.js';

import { chartDataPointStyleRole, chartSeriesSourceIndex } from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';

import { resolveDataLabelPlacement, type DataLabelRect } from '../data-label-layout.js';

import { EMU_PER_PT } from '../../units.js';

import {
  indexPointOverrides,
  pieSliceColor,
  piePointStyleIndex,
  paintClassicPiePointOutline,
  chartFontFamily,
  drawLegendSwatch,
  DataLabelLegendKey,
  createDataLabelLegendKeyResolver,
  LEGEND_SWATCH_TEXT_GAP,
  legendSwatchWidths,
  legendSwatchHeight,
  measuredLegendReserve,
  drawLegendForLayout,
  drawChartTitleForLayout,
  dataLabelRectIntersection,
  applyDecorationLineStyle,
  chartStyleRoleLine,
  chartStyleRoleLeaderLine,
  customRichDataLabelOptions,
  drawBoundedDataLabelText,
  dashPatternForPreset,
  paintClassicDataPointPath,
  paintClassicDataPointRect,
} from '../shared/classic.js';

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut — supports dataPointColors (per slice).
// ═══════════════════════════════════════════════════════════════════════════

/** Inside-radius fraction (of the outer radius) for a SOLID pie's `ctr` / `inEnd`
 *  / `bestFit` data labels (§21.2.2.48). PowerPoint places these near the rim,
 *  not at the disc mid-radius: four observed slice sizes place labels at
 *  0.878 / 0.888 / 0.887 / 0.912·outerR — a flat near-rim
 *  constant independent of slice angle (see the `labelR` comment in
 *  {@link drawPieRichLabels}). 0.88 is the empirical fit; it is an approximation
 *  of an undocumented PowerPoint layout, not a spec-defined geometry. Doughnut
 *  labels use the exact ring midpoint instead and never consult this. */
const PIE_CTR_LABEL_RADIUS_FRAC = 0.88;

interface OfPiePoint {
  sourceIndex: number;
  value: number;
}

/** ECMA-376 §21.2.2.126 pie-of-pie / bar-of-pie. Source point identity stays
 * attached to every detail item; only the primary plot receives one aggregate
 * slice. Automatic geometry is deliberately compact and parameterized solely
 * by the authored gap and second-plot size. */
export function renderOfPieChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const source = chart.series[0];
  if (!source) return;
  const secondarySet = planOfPieSecondaryIndices(chart.ofPie, source.values);
  if (secondarySet == null || secondarySet.size === 0) {
    renderPieChart(
      ctx, { ...chart, chartType: 'pie' }, r, false, ptToPx, shapeRotationDeg,
    );
    return;
  }
  const primary: OfPiePoint[] = [];
  const secondary: OfPiePoint[] = [];
  for (let sourceIndex = 0; sourceIndex < source.values.length; sourceIndex++) {
    const sourceValue = source.values[sourceIndex];
    const value = sourceValue == null ? 0 : Math.abs(sourceValue);
    if (!(value > 0) || !Number.isFinite(value)) continue;
    (secondarySet.has(sourceIndex) ? secondary : primary).push({ sourceIndex, value });
  }
  if (secondary.length === 0) {
    renderPieChart(
      ctx, { ...chart, chartType: 'pie' }, r, false, ptToPx, shapeRotationDeg,
    );
    return;
  }

  const legendChart: ChartModel = { ...chart, chartType: 'pie' };
  const legend = measuredLegendReserve(ctx, legendChart, r.w, r.h, 0.28, ptToPx);
  const frame = computeChartFrame(chart, r.x, r.y, r.w, r.h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.28,
    legendReserve: legend,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(
    ctx, chart, r.x, r.y, r.w, r.h,
    r.y + frame.title.topPad, frame.title.fontPx,
  );
  const { px0, py0, pw, ph } = frame.plotRect;
  if (!(pw > 0) || !(ph > 0)) return;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const options = chart.ofPie;
  const sizeRatio = Math.max(0.05, Math.min(2, (options?.secondPieSizePercent ?? 75) / 100));
  const gapUnits = Math.max(0, options?.gapWidthPercent ?? 150) / 100;
  const mainRadius = Math.min(ph * 0.44, (pw * 0.9) / (2 + 2 * sizeRatio + gapUnits));
  const detailRadius = mainRadius * sizeRatio;
  if (!(mainRadius > 0) || !(detailRadius > 0)) return;
  const usedW = 2 * mainRadius + gapUnits * mainRadius + 2 * detailRadius;
  const left = px0 + (pw - usedW) / 2;
  const mainCx = left + mainRadius;
  const detailCx = left + 2 * mainRadius + gapUnits * mainRadius + detailRadius;
  const cy = py0 + ph / 2;
  const secondaryTotal = secondary.reduce((sum, point) => sum + point.value, 0);
  const pointOverrides = indexPointOverrides(source.dataPointOverrides);
  const mainPoints: OfPiePoint[] = [
    ...primary,
    { sourceIndex: secondary[0].sourceIndex, value: secondaryTotal },
  ];

  const drawPie = (
    points: OfPiePoint[], cx: number, radius: number,
  ): { aggregateStart: number; aggregateEnd: number } => {
    const total = points.reduce((sum, point) => sum + point.value, 0);
    let angle = -Math.PI / 2;
    let aggregateStart = angle;
    let aggregateEnd = angle;
    for (let index = 0; index < points.length; index++) {
      const point = points[index];
      const sweep = total > 0 ? point.value / total * Math.PI * 2 : 0;
      const pointOverride = pointOverrides.get(point.sourceIndex);
      const styleIndex = piePointStyleIndex(chart, source, 0, point.sourceIndex);
      const fallbackColor = pieSliceColor(
        point.sourceIndex, source, chart.varyColors !== false, 0,
      );
      const fillDecision = classicDataPointFillDecision(
        chart, source, pointOverride, styleIndex, point.sourceIndex,
      );
      const sliceStart = angle;
      const sliceEnd = angle + sweep;
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(pointOverride?.chartexStyle, source.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPoint', 0),
        styleIndex,
        { x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2 },
        ptToPx,
        target => {
          target.beginPath();
          target.moveTo(cx, cy);
          target.arc(cx, cy, radius, sliceStart, sliceEnd);
          target.closePath();
          paintClassicDataPointPath(target, fillDecision, {
            x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2,
          }, fallbackColor, ptToPx, shapeRotationDeg);
          paintClassicPiePointOutline(
            target, chart, source, pointOverride, styleIndex,
            '#FFFFFF', ptToPx,
            { x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2 },
            shapeRotationDeg,
          );
        },
      );
      if (index === points.length - 1) {
        aggregateStart = angle;
        aggregateEnd = angle + sweep;
      }
      angle += sweep;
    }
    return { aggregateStart, aggregateEnd };
  };

  const aggregateAngles = drawPie(mainPoints, mainCx, mainRadius);
  let connectorTop = cy - detailRadius;
  let connectorBottom = cy + detailRadius;
  if ((options?.type ?? 'pie') === 'bar') {
    let top = cy - detailRadius;
    const barW = detailRadius;
    for (const point of secondary) {
      const height = secondaryTotal > 0 ? 2 * detailRadius * point.value / secondaryTotal : 0;
      const bx = detailCx - barW / 2;
      const pointOverride = pointOverrides.get(point.sourceIndex);
      const styleIndex = piePointStyleIndex(chart, source, 0, point.sourceIndex);
      const fallbackColor = pieSliceColor(
        point.sourceIndex, source, chart.varyColors !== false, 0,
      );
      const fillDecision = classicDataPointFillDecision(
        chart, source, pointOverride, styleIndex, point.sourceIndex,
      );
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(pointOverride?.chartexStyle, source.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPoint', 0),
        styleIndex,
        { x: bx, y: top, w: barW, h: height },
        ptToPx,
        target => {
          paintClassicDataPointRect(
            target, fillDecision, { x: bx, y: top, w: barW, h: height },
            fallbackColor, ptToPx, shapeRotationDeg,
          );
          target.beginPath();
          target.rect(bx, top, barW, height);
          paintClassicPiePointOutline(
            target, chart, source, pointOverride, styleIndex,
            '#FFFFFF', ptToPx, { x: bx, y: top, w: barW, h: height },
            shapeRotationDeg,
          );
        },
      );
      top += height;
    }
    connectorTop = cy - detailRadius;
    connectorBottom = cy + detailRadius;
  } else {
    drawPie(secondary, detailCx, detailRadius);
  }

  if (options?.seriesLines ?? true) {
    const resolvedSeriesLine = chartStyleRoleLine(
      chart,
      options?.seriesLineStyle ?? {},
      'seriesLine',
    );
    const paintsSeriesLine = applyDecorationLineStyle(ctx, {
      ...resolvedSeriesLine,
      color: resolvedSeriesLine.color ?? '808080',
    }, ptToPx);
    const fromTop = {
      x: mainCx + Math.cos(aggregateAngles.aggregateStart) * mainRadius,
      y: cy + Math.sin(aggregateAngles.aggregateStart) * mainRadius,
    };
    const fromBottom = {
      x: mainCx + Math.cos(aggregateAngles.aggregateEnd) * mainRadius,
      y: cy + Math.sin(aggregateAngles.aggregateEnd) * mainRadius,
    };
    if (paintsSeriesLine) {
      ctx.beginPath(); ctx.moveTo(fromTop.x, fromTop.y); ctx.lineTo(detailCx - detailRadius, connectorTop); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(fromBottom.x, fromBottom.y); ctx.lineTo(detailCx - detailRadius, connectorBottom); ctx.stroke();
    }
  }
  if (legend) {
    drawLegendForLayout(
      ctx, legendChart, legend,
      r.x, r.y, r.w, r.h, px0, py0, pw, ph, frame.title.bandH + 2, ptToPx,
    );
  }
}

export function renderPieChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  isDoughnut: boolean,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const s = chart.series[0]; if (!s) return;
  const cats = (s.categories && s.categories.length > 0) ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;
  const seriesLegend = isDoughnut
    && chart.varyColors === false
    && chart.series.length > 1;
  const legendChart: ChartModel = seriesLegend
    ? chart
    : { ...chart, series: [{ ...s, categories: cats }] };

  // Shared frame (radial form). Pie uses title pads 0.035 / 0.035; its legend
  // labels categories (one row per slice) so it reserves a wider 0.28 side band
  // (vs the default 0.22). The h*0.02 gap below the title/legend before centring
  // is the shared radial gap. Params keep pixels unchanged.
  const pieLeg = measuredLegendReserve(ctx, legendChart, w, h, 0.28, ptToPx);
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.28,
    legendReserve: pieLeg,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  const titleFontPx = frame.title.fontPx;
  const titleH = frame.title.bandH;
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, titleFontPx);

  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(
    ctx, chart, plotLeft, plotTop, pw, ph, ptToPx, shapeRotationDeg,
  );
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  const outerR = Math.min(pw, ph) * 0.42;

  // §21.2.2.52 firstSliceAng: the first slice begins `firstSliceAngle` degrees
  // clockwise from 12 o'clock. Canvas 0 rad points right (+x) and its angles
  // grow clockwise (y-down), so 12 o'clock is −90°. Default 0 keeps the
  // historical −90° start (byte-stable for files without the element).
  const startAngle = -Math.PI / 2 + ((chart.firstSliceAngle ?? 0) * Math.PI) / 180;

  // §21.2.2.60 holeSize (doughnut only): hole diameter as 1–90% of the outer
  // diameter. The ECMA schema default is 10%, but a real doughnut always writes
  // an explicit holeSize (Office emits 50–75%); 50% is the historical inner
  // radius, so an absent holeSize keeps the prior look (byte-stable). Pie has
  // no hole (innerR = 0).
  const holePct = isDoughnut ? Math.max(1, Math.min(90, chart.holeSize ?? 50)) : 0;

  // Concentric rings. Doughnut plots EVERY series as a ring (outermost =
  // series[0]); pie plots only series[0]. The band from the hole radius to the
  // outer radius is split evenly across the rings. A single-series doughnut is
  // byte-identical to the prior single-ring geometry.
  const rings = isDoughnut ? chart.series : [s];
  const pointOverridesBySeries = new Map(
    rings.map(series => [series, indexPointOverrides(series.dataPointOverrides)] as const),
  );
  const innerR = outerR * (holePct / 100);
  const ringBand = (outerR - innerR) / rings.length;

  // Explosion offset for slice `i` of series `ser`: move the slice out from the
  // center along its mid-angle by `explosion`% of the outer radius. §21.2.2.61
  // only defines `explosion` as an unbounded `xsd:unsignedInt` "amount the data
  // point shall be moved from the center of the pie" — the 0-100-as-percent
  // interpretation is a de-facto Office convention (the Point Explosion UI
  // slider), not a spec-mandated range (see `ChartDataPointOverride.explosion`
  // in types/chart.ts). Absent / zero explosion → no offset (byte-stable).
  const explodeOffset = (ser: ChartSeries, i: number): number => {
    const e = pointOverridesBySeries.get(ser)?.get(i)?.explosion ?? ser.explosion ?? 0;
    return e > 0 ? (e / 100) * outerR : 0;
  };

  // The legacy `showDataLabels` percent label (drawn INLINE per slice on the
  // outer ring, exactly as before) is used only when the series has no rich
  // `<c:dLbls>` definition; the rich labels are drawn in a separate pass after
  // all slices. Keeping the legacy path inline preserves the historical
  // draw-call order for a plain pie/doughnut (byte-stable).
  const richDef: ChartSeriesDataLabels = s.seriesDataLabels ?? {
    showVal: false,
    showCatName: false,
    showSerName: false,
    showPercent: false,
  };
  // A point-level dLbl is independently authored and must not depend on a
  // series dLbls default existing. Even a delete-only override participates in
  // rich dispatch so the legacy chart-wide percent path cannot resurrect it.
  const hasRichLabels = s.seriesDataLabels != null || (s.dataLabelOverrides?.length ?? 0) > 0;
  const legacyLabels = chart.showDataLabels && !hasRichLabels;
  const dLblFont = chartFontFamily(
    chart, richDef.fontFace ?? chart.dataLabelFontFace, 'minor',
  );

  for (let ring = 0; ring < rings.length; ring++) {
    const rs = rings[ring];
    const rVals = rs.values.map(v => Math.abs(v ?? 0));
    const rTotal = rVals.reduce((a, b) => a + b, 0);
    if (rTotal === 0) continue;
    // Ring 0 is the OUTERMOST band; deeper rings step inward toward the hole.
    const rOuter = outerR - ring * ringBand;
    const rInner = rOuter - ringBand;

    let angle = startAngle;
    for (let i = 0; i < rVals.length; i++) {
      const slice = (rVals[i] / rTotal) * Math.PI * 2;
      // A zero-valued wedge has no path area, label, or angular advance. Skip
      // before resolving structured paint/effects so empty points cannot force
      // thousands of gradient stops, image tiles, or auxiliary rasters.
      if (!(slice > 0)) continue;
      const styleIndex = piePointStyleIndex(chart, rs, ring, i);
      const color = pieSliceColor(i, rs, chart.varyColors !== false, ring);
      const point = pointOverridesBySeries.get(rs)?.get(i);
      const midAngle = angle + slice / 2;
      const off = explodeOffset(rs, i);
      const ox = off > 0 ? Math.cos(midAngle) * off : 0;
      const oy = off > 0 ? Math.sin(midAngle) * off : 0;
      const fillDecision = classicDataPointFillDecision(chart, rs, point, styleIndex, i);
      const sliceBounds = {
        x: cx2 + ox - rOuter,
        y: cy2 + oy - rOuter,
        w: rOuter * 2,
        h: rOuter * 2,
      };
      const paintSlice = (target: CanvasRenderingContext2D): void => {
        target.beginPath();
        if (rInner > 0.01) {
          // Annular slice (doughnut ring): outer arc CW, inner arc CCW.
          target.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
          target.arc(cx2 + ox, cy2 + oy, rInner, angle + slice, angle, true);
        } else {
          // Solid wedge (pie, or the innermost pie-like ring).
          target.moveTo(cx2 + ox, cy2 + oy);
          target.arc(cx2 + ox, cy2 + oy, rOuter, angle, angle + slice);
        }
        target.closePath();
        paintClassicDataPointPath(
          target, fillDecision, sliceBounds, color, ptToPx, shapeRotationDeg,
        );
        paintClassicPiePointOutline(
          target, chart, rs, point, styleIndex, color, ptToPx, sliceBounds, shapeRotationDeg,
        );
      };
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(point?.chartexStyle, rs.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPoint', chartSeriesSourceIndex(chart, rs)),
        styleIndex,
        sliceBounds,
        ptToPx,
        paintSlice,
      );

      // Legacy percent label — outer ring only, drawn inline (byte-stable).
      if (legacyLabels && ring === 0 && slice > 0.15) {
        const labelR = outerR * (isDoughnut ? 0.75 : 0.6);
        const lx2 = cx2 + ox + Math.cos(midAngle) * labelR;
        const ly2 = cy2 + oy + Math.sin(midAngle) * labelR;
        const pct2 = Math.round((rVals[i] / rTotal) * 100);
        const lsz = Math.max(8, outerR * 0.1);
        ctx.font = `bold ${lsz}px ${dLblFont}`;
        ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(`${pct2}%`, lx2, ly2);
      }

      angle += slice;
    }
  }

  // Rich data labels (`<c:dLbls>`: showVal / showCatName / showSerName /
  // showPercent + dLblPos, §21.2.2.35), drawn on the OUTER ring after all
  // slices. Only runs when a rich definition is present; the plain percent
  // labels above are byte-identical to the pre-CH8 pie.
  if (hasRichLabels) {
    const outerRingInnerR = isDoughnut ? outerR - ringBand : 0;
    drawPieRichLabels(
      ctx, chart, richDef, s, cats, vals, total,
      cx2, cy2, outerR, outerRingInnerR, startAngle, dLblFont, ptToPx,
      plotLeft, plotTop, pw, ph,
      x, y, w, h,
      shapeRotationDeg,
    );
  }

  if (pieLeg) {
    // Varying pie/doughnut legends are category-driven and match each slice.
    // An explicitly non-varying multi-series doughnut is series-driven instead
    // and identifies its solid-colour rings, matching Excel-produced output.
    drawLegendForLayout(
      ctx, legendChart, pieLeg,
      x, y, w, h, plotLeft, plotTop, pw, ph, titleH + 2,
      ptToPx,
    );
  }
}

/** Draw the rich outer-ring data labels for a pie / doughnut from a series-level
 *  `<c:dLbls>` (§21.2.2.35: showVal / showCatName / showSerName / showPercent +
 *  dLblPos). Only called when such a definition exists; the plain percent-label
 *  path stays inline in the slice loop (byte-stable). `font` is the pre-resolved
 *  data-label CSS font-family.
 *
 *  When the `<c:dLbls>` carries a callout-box shape (`<c:spPr>` → `def.labelBox`,
 *  §21.2.2.197) the labels are drawn Word-style: each is a boxed callout placed
 *  OUTSIDE its slice at the slice mid-angle, with adjacent boxes pushed apart to
 *  avoid overlap (`bestFit`), and a leader line back to the rim for any box that
 *  ends up far from its slice. Plain `outEnd` labels use the same outside-rim
 *  invariant without painting a box; the inside positions retain their radial
 *  layout. */
function drawPieRichLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number, innerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  plotX: number, plotY: number, plotW: number, plotH: number,
  chartX: number, chartY: number, chartW: number, chartH: number,
  shapeRotationDeg: number,
): void {
  const overrides = s.dataLabelOverrides ?? [];
  const overridesByIndex = indexPointOverrides(overrides);
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(
    { ...chart, series: [{ ...s, categories: cats }] },
    ptToPx,
  );
  const calloutIndices = new Set<number>();
  for (let index = 0; index < vals.length; index++) {
    if (s.sourceHidden?.[index] === true) continue;
    const override = overridesByIndex.get(index);
    if (dataLabelIsDeleted(def, override)) continue;
    const labelBox = mergeChartLabelBoxes(override?.labelBox, def.labelBox);
    // A border-only label shape is an outline around the ordinary radial
    // label; it does not opt into Word's filled boxed-callout layout. Treating
    // any visible border as a callout invented leaders for Excel bestFit pie
    // labels even though the authored labels remain on their slices.
    const hasVisibleCalloutFill = labelBox?.fillHidden !== true
      && (labelBox?.fill != null || labelBox?.fillPaint != null);
    if (hasVisibleCalloutFill) {
      calloutIndices.add(index);
    }
  }
  // Boxed labels have their own paint/collision pass, but only those points are
  // dispatched to it. A per-point spPr must not turn every sibling into a
  // callout or change the sibling's authored radial position.
  if (calloutIndices.size > 0) {
    drawPieCalloutLabels(
      ctx, chart, def, s, cats, vals, total, cx2, cy2, outerR, innerR, startAngle,
      font, ptToPx, plotX, plotW, plotY, plotH, chartX, chartY, chartW, chartH,
      calloutIndices, overridesByIndex, shapeRotationDeg,
    );
  }

  const outsideLabels: PieOutsideLabel[] = [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    if (s.sourceHidden?.[i] === true) continue;
    if (calloutIndices.has(i)) continue;
    // A per-point `<c:dLbl idx>` (§21.2.2.47) overrides the series-level
    // `<c:dLbls>` (§21.2.2.49) for this one slice. Its show-flags, font color /
    // size / bold, and position each fall back to the series default when the
    // point declares none. A point may set `showCatName=0 showPercent=1` plus
    // white text while the series default is
    // `showCatName=1` black — so honoring the per-point flags is what makes the
    // labels render as white percent-only (matching PowerPoint / the PDF).
    const ov = overridesByIndex.get(i);
    // A genuinely deleted label (`<c:delete val="1">`, §21.2.2.43) is skipped.
    // A style/flag-only `<c:dLbl>` (no `<c:tx>`) is NOT a delete. Such slices
    // can carry `text: ""` with white/percent-only flag overrides, so we
    // key off the explicit `deleted` flag, never the empty text.
    if (dataLabelIsDeleted(def, ov)) continue;
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    const showLegendKey = ov?.showLegendKey ?? def.showLegendKey ?? false;
    // §21.2.2.35 label composition. A per-point custom `<c:tx>` (non-empty
    // override text) wins outright; otherwise compose from the resolved flags.
    // Positioning is handled below (§21.2.2.48 `dLblPos`). Percent is the
    // slice's share of the total.
    const text = effectiveDataLabelText({
      customText: ov?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: (cats[i] ?? '').toString(),
      seriesName: s.name,
      sourceValue: vals[i],
      percentRatio: vals[i] / total,
      formatCode: ov?.formatCode ?? def.formatCode ?? s.valFormatCode ?? null,
      percentFormatCode: ov?.formatCode ?? def.formatCode ?? '0%',
      date1904: chart.date1904 ?? false,
      separator: ov?.separator ?? def.separator,
    });
    const legendKey = showLegendKey ? dataLabelLegendKey(0, i) : undefined;
    if (!text && !legendKey) continue;
    const pos = ov?.position ?? def.position ?? 'bestFit';
    const outside = pos === 'outEnd';
    const sizeHpt = ov?.fontSizeHpt ?? def.fontSizeHpt;
    const sizePx = chartTextFontSizePx(sizeHpt, ptToPx) ?? Math.max(8, outerR * 0.1);
    const bold = ov?.fontBold ?? def.fontBold;
    const fontColor = ov?.fontColor ?? def.fontColor;
    const labelFont = (ov?.fontFace ?? def.fontFace)
      ? chartFontFamily(chart, ov?.fontFace ?? def.fontFace, 'minor')
      : font;
    const textStyle = effectiveDataLabelTextStyle(ov, def);
    const rich = customRichDataLabelOptions(
      chart, ov, ptToPx, labelFont, bold ?? false, textStyle,
    );
    const automaticLabelR = innerR > 0.01
      ? (innerR + outerR) / 2
      : outerR * PIE_CTR_LABEL_RADIUS_FRAC;
    if (ov?.manualLayout) {
      ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
      drawBoundedDataLabelText(
        ctx,
        text,
        {
          kind: 'point',
          x: cx2 + Math.cos(midAngle) * automaticLabelR,
          y: cy2 + Math.sin(midAngle) * automaticLabelR,
          position: 'ctr',
        },
        { x: plotX, y: plotY, w: plotW, h: plotH },
        sizePx,
        fontColor ? `#${fontColor}` : '#fff',
        ov.manualLayout,
        { x: chartX, y: chartY, w: chartW, h: chartH },
        rich,
        legendKey,
        textStyle,
        ptToPx,
        mergeChartLabelBoxes(ov?.labelBox, def.labelBox),
        shapeRotationDeg,
      );
      continue;
    }
    if (outside) {
      ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
      const richBlock = rich
        ? resolveRichDataLabelBlock(ctx, rich, sizePx, fontColor ? `#${fontColor}` : '#333')
        : null;
      const lineHeight = sizePx * 1.15;
      const fittedLines = richBlock ? [] : fitStyledDataLabelLines(
        text, Math.max(0, chartW - sizePx), Math.max(0, chartH - sizePx),
        lineHeight, value => ctx.measureText(value).width, textStyle,
      );
      if (rich && !richBlock) continue;
      if (!richBlock && fittedLines.length === 0 && !legendKey) continue;
      const textW = richBlock?.width ?? fittedLines.reduce(
        (max, line) => Math.max(max, ctx.measureText(line).width), 0,
      );
      const textH = richBlock?.height
        ?? (sizePx + Math.max(0, fittedLines.length - 1) * lineHeight);
      const keyW = legendKey
        ? (legendSwatchWidths([legendKey.entry], sizePx, ptToPx)[0] ?? 0)
        : 0;
      const keyH = legendKey ? legendSwatchHeight(legendKey.entry, sizePx, ptToPx) : 0;
      outsideLabels.push(createPieOutsideLabel(
        fittedLines, midAngle, cx2, cy2, outerR,
        Math.min(keyW + (text ? LEGEND_SWATCH_TEXT_GAP : 0) + textW, Math.max(0, chartW - sizePx)),
        Math.min(Math.max(keyH, textH), Math.max(0, chartH - sizePx)),
        lineHeight, sizePx, bold ?? false,
        fontColor ? `#${fontColor}` : '#333',
        labelFont,
        richBlock ?? undefined,
        legendKey,
        textStyle,
        ptToPx,
      ));
      continue;
    }
    // §21.2.2.48 ST_DLblPos radial placement. The spec enumerates the positions
    // (bestFit / ctr / inEnd / outEnd …) but gives no geometry, so the inside
    // radii below reproduce the bounded Office vector observations for solid
    // pie and doughnut labels:
    //
    //   • DOUGHNUT (innerR > 0), ctr / inEnd / bestFit → the RING midpoint
    //     (innerR + outerR)/2. Verified on the 55%-hole doughnut: labels sit at
    //     0.772–0.778·outerR ≈ (0.55+1)/2 = 0.775. Byte-stable — unchanged.
    //   • SOLID pie (innerR ≈ 0), ctr / inEnd / bestFit → ≈0.88·outerR, NOT the
    //     disc mid-radius. Measured label-centroid ratios across the 54/27/14/5%
    //     slices were 0.878 / 0.888 / 0.887 / 0.912 (center + outer radius from a
    //     least-squares rim fit, residual std 0.43pt), i.e. a flat near-rim
    //     constant independent of slice angle — so it is a fixed fraction, not a
    //     sector centroid. The 5% sliver rides marginally further out in
    //     PowerPoint; we do not model that per-slice nudge. This is an empirical
    //     approximation of an undocumented PowerPoint layout, not a spec formula.
    const labelR = automaticLabelR;
    const lx2 = cx2 + Math.cos(midAngle) * labelR;
    const ly2 = cy2 + Math.sin(midAngle) * labelR;
    ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${labelFont}`;
    const tangentialCapacity = 2 * labelR * Math.sin(Math.min(Math.PI, Math.abs(slice)) / 2)
      - sizePx;
    const radialCapacity = innerR > 0.01
      ? outerR - innerR - sizePx
      : outerR - sizePx;
    if (!(tangentialCapacity > 0) || !(radialCapacity > 0)) continue;
    const sliceBounds = dataLabelRectIntersection(
      {
        x: lx2 - tangentialCapacity / 2,
        y: ly2 - radialCapacity / 2,
        w: tangentialCapacity,
        h: radialCapacity,
      },
      { x: plotX, y: plotY, w: plotW, h: plotH },
    );
    if (!sliceBounds) continue;
    drawBoundedDataLabelText(
      ctx,
      text,
      { kind: 'point', x: lx2, y: ly2, position: 'ctr' },
      sliceBounds,
      sizePx,
      fontColor ? `#${fontColor}` : '#fff',
      undefined,
      { x: chartX, y: chartY, w: chartW, h: chartH },
      rich,
      legendKey,
      textStyle,
      ptToPx,
      mergeChartLabelBoxes(ov?.labelBox, def.labelBox),
      shapeRotationDeg,
    );
  }

  drawPieOutsideLabels(ctx, outsideLabels, chartX, chartY, chartW, chartH);
}

/** Plain `<c:dLblPos val="outEnd">` label block. */
interface PieOutsideLabel {
  lines: string[];
  rich?: RichDataLabelBlock;
  legendKey?: DataLabelLegendKey;
  boxW: number;
  boxH: number;
  unrotatedW: number;
  unrotatedH: number;
  textStyle: DataLabelTextStyle;
  ptToPx: number;
  lineHeight: number;
  fontPx: number;
  bold: boolean;
  fontColor: string;
  font: string;
  cxBox: number;
  cyBox: number;
}

function pointToRectDistance(
  px: number, py: number,
  rectCx: number, rectCy: number,
  halfW: number, halfH: number,
): number {
  const dx = Math.max(Math.abs(rectCx - px) - halfW, 0);
  const dy = Math.max(Math.abs(rectCy - py) - halfH, 0);
  return Math.hypot(dx, dy);
}

/** Find the first point on a slice-midpoint ray whose complete visible label
 * rectangle clears the pie. This restores the release geometry without
 * reintroducing collision moves or their synthetic leader lines. */
function outsideLabelRadialDistance(
  midAngle: number,
  outerR: number,
  halfW: number,
  halfH: number,
  clearance: number,
): number {
  const ux = Math.cos(midAngle);
  const uy = Math.sin(midAngle);
  const target = outerR + clearance;
  let low = 0;
  let high = target + Math.hypot(halfW, halfH);
  for (let i = 0; i < 32; i++) {
    const mid = (low + high) / 2;
    const distance = pointToRectDistance(0, 0, ux * mid, uy * mid, halfW, halfH);
    if (distance >= target) high = mid;
    else low = mid;
  }
  return high;
}

function createPieOutsideLabel(
  lines: string[],
  midAngle: number,
  pieCx: number,
  pieCy: number,
  outerR: number,
  boxW: number,
  boxH: number,
  lineHeight: number,
  fontPx: number,
  bold: boolean,
  fontColor: string,
  font: string,
  rich?: RichDataLabelBlock,
  legendKey?: DataLabelLegendKey,
  textStyle: DataLabelTextStyle = {},
  ptToPx = 1,
): PieOutsideLabel {
  const visibleRotated = rotatedDataLabelSize(
    boxW, boxH, textStyle.textRotation, textStyle.textVerticalMode,
  );
  const insets = dataLabelInsets(textStyle, ptToPx);
  const unrotatedW = boxW + insets.left + insets.right;
  const unrotatedH = boxH + insets.top + insets.bottom;
  const rotated = rotatedDataLabelSize(
    unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
  );
  boxW = rotated.w;
  boxH = rotated.h;
  // `outEnd` requires the visible label content, rather than only its anchor,
  // to clear the pie. Keep the release-era radial geometry while leaving each
  // label on its authored slice-midpoint ray; no collision movement means no
  // synthetic leader line is introduced.
  const distance = outsideLabelRadialDistance(
    midAngle, outerR, visibleRotated.w / 2, visibleRotated.h / 2, fontPx * 0.5,
  );
  const cxBox = pieCx + Math.cos(midAngle) * distance;
  const cyBox = pieCy + Math.sin(midAngle) * distance;
  return {
    lines, rich, legendKey,
    boxW, boxH, unrotatedW, unrotatedH, textStyle, ptToPx,
    lineHeight, fontPx, bold, fontColor, font,
    cxBox, cyBox,
  };
}

/** Paint automatic plain outEnd labels at their authored slice-midpoint anchors.
 * Rich callout labels retain their separate bounded collision/leader resolver. */
function drawPieOutsideLabels(
  ctx: CanvasRenderingContext2D,
  labels: PieOutsideLabel[],
  boundsX: number,
  boundsY: number,
  boundsW: number,
  boundsH: number,
): void {
  if (labels.length === 0) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(boundsX, boundsY, boundsW, boundsH);
  ctx.clip();

  for (const label of labels) {
    const insets = dataLabelInsets(label.textStyle, label.ptToPx);
    const rotated = rotatedDataLabelSize(
      label.unrotatedW, label.unrotatedH,
      label.textStyle.textRotation, label.textStyle.textVerticalMode,
    );
    const textCx = label.cxBox + (insets.left - insets.right) / 2;
    const textCy = label.cyBox + (insets.top - insets.bottom) / 2;
    const innerWidth = Math.max(0, label.unrotatedW - insets.left - insets.right);
    const paintAlign = dataLabelCanvasTextAlign(label.textStyle, 'center');
    const textAnchorX = paintAlign === 'left'
      ? label.cxBox - label.unrotatedW / 2 + insets.left
      : paintAlign === 'right'
        ? label.cxBox + label.unrotatedW / 2 - insets.right
        : textCx;
    ctx.save();
    if (rotated.radians !== 0) {
      ctx.translate(label.cxBox, label.cyBox);
      ctx.rotate(rotated.radians);
      ctx.translate(-label.cxBox, -label.cyBox);
    }
    if (!label.legendKey) {
      if (label.rich) {
        paintRichDataLabelBlock(
          ctx, label.rich, textAnchorX, textCy, paintAlign, 'middle', innerWidth,
        );
        ctx.restore();
        continue;
      }
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.bold ? 'bold ' : ''}${label.fontPx}px ${label.font}`;
      ctx.fillStyle = label.fontColor;
      ctx.textAlign = paintAlign;
      ctx.textBaseline = 'middle';
      const baselineShift = (label.textStyle.fontBaseline ?? 0) * label.fontPx;
      const firstY = textCy - ((label.lines.length - 1) * label.lineHeight) / 2 - baselineShift;
      if (!(label.textStyle.fontPaintAuthored === true
        && (label.textStyle.fontHidden === true || label.textStyle.fontColor == null))) {
        for (let i = 0; i < label.lines.length; i++) {
          ctx.fillText(label.lines[i], textAnchorX, firstY + i * label.lineHeight);
        }
      }
      ctx.restore();
      continue;
    }
    ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.bold ? 'bold ' : ''}${label.fontPx}px ${label.font}`;
    const keyWidth = label.legendKey
      ? (legendSwatchWidths([label.legendKey.entry], label.fontPx, label.legendKey.ptToPx)[0] ?? 0)
      : 0;
    const keyHeight = label.legendKey
      ? legendSwatchHeight(label.legendKey.entry, label.fontPx, label.legendKey.ptToPx)
      : 0;
    const textWidth = label.rich?.width ?? label.lines.reduce(
      (max, line) => Math.max(max, ctx.measureText(line).width), 0,
    );
    const gap = label.legendKey && (label.rich || label.lines.length > 0)
      ? LEGEND_SWATCH_TEXT_GAP
      : 0;
    const contentWidth = keyWidth + gap + textWidth;
    const contentLeft = textCx - contentWidth / 2;
    if (label.legendKey) {
      drawLegendSwatch(
        ctx,
        label.legendKey.entry.swatchStyle,
        label.legendKey.entry.color,
        contentLeft,
        textCy - keyHeight / 2,
        keyWidth,
        keyHeight,
        label.legendKey.entry.marker,
        label.legendKey.entry.fillPaint,
        label.legendKey.entry.outlinePaint,
        label.legendKey.entry.outlineColor,
        label.legendKey.entry.outlineWidthEmu,
        label.legendKey.entry.outlineDash,
        label.legendKey.entry.outlineCustomDash,
        label.legendKey.entry.outlineCap,
        label.legendKey.entry.outlineJoin,
        label.legendKey.ptToPx,
        label.legendKey.shapeRotationDeg,
        label.legendKey.entry.directEffect,
        label.legendKey.entry.fallbackEffect,
        label.legendKey.entry.directEffectIndex,
        label.legendKey.entry.fallbackEffectIndex,
      );
    }
    if (label.rich) {
      paintRichDataLabelBlock(
        ctx, label.rich, contentLeft + keyWidth + gap, textCy, 'left', 'middle',
      );
      ctx.restore();
      continue;
    }
    ctx.fillStyle = label.fontColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const baselineShift = (label.textStyle.fontBaseline ?? 0) * label.fontPx;
    const firstY = textCy - ((label.lines.length - 1) * label.lineHeight) / 2 - baselineShift;
    if (!(label.textStyle.fontPaintAuthored === true
      && (label.textStyle.fontHidden === true || label.textStyle.fontColor == null))) {
      for (let i = 0; i < label.lines.length; i++) {
        ctx.fillText(label.lines[i], contentLeft + keyWidth + gap, firstY + i * label.lineHeight);
      }
    }
    ctx.restore();
  }
  ctx.restore();
}

/** One laid-out pie callout label: its wrapped text lines, box rectangle, the
 *  rim anchor point on its slice, and the resolved per-point style. */
interface PieCalloutLabel {
  lines: string[];
  rich?: RichDataLabelBlock;
  legendKey?: DataLabelLegendKey;
  lineHeight: number;
  /** Slice mid-angle (canvas radians) — the leader-line target direction. */
  midAngle: number;
  /** Rim anchor point (on the outer arc at `midAngle`). */
  rimX: number;
  rimY: number;
  /** Half-height of the text block (px) — box grows symmetrically around cy. */
  boxW: number;
  boxH: number;
  unrotatedW: number;
  unrotatedH: number;
  /** Box centre (mutated by the collision pass). */
  cxBox: number;
  cyBox: number;
  /** true when the label sits on the left half (box hangs to the left). */
  leftSide: boolean;
  fontColor: string;
  box?: ChartLabelBox;
  fontPx: number;
  bold: boolean;
  font: string;
  textStyle: DataLabelTextStyle;
  ptToPx: number;
  /** An authored inside position keeps the box at its slice anchor. */
  inside: boolean;
  /** Explicit per-point manual layout is never moved by the auto collision pass. */
  manualClip?: DataLabelRect;
}

/** Word-style boxed pie/doughnut callout labels (`bestFit`). Each label is a
 *  filled+bordered rectangle placed just outside its slice at the slice
 *  mid-angle; adjacent boxes on the same side are pushed vertically apart so
 *  they do not overlap, and a leader line is drawn back to the rim for any box
 *  whose gap from the rim exceeds a small threshold. Style (box fill/border,
 *  leader colour/width, per-point font colour and box overrides) all comes from
 *  the parsed model — no empirical constants beyond the layout paddings, which
 *  are geometry (not spec values). */
function drawPieCalloutLabels(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  def: ChartSeriesDataLabels,
  s: ChartSeries,
  cats: string[],
  vals: number[],
  total: number,
  cx2: number, cy2: number,
  outerR: number, innerR: number,
  startAngle: number,
  font: string,
  ptToPx: number,
  boundsX: number, boundsW: number, boundsY: number, boundsH: number,
  chartX: number, chartY: number, chartW: number, chartH: number,
  indices: ReadonlySet<number>,
  overridesByIndex: ReadonlyMap<number, ChartDataLabelOverride>,
  shapeRotationDeg: number,
): void {
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  const findOverride = (i: number): ChartDataLabelOverride | undefined =>
    overridesByIndex.get(i);

  // Base font size: series default (hpt → px) or a radius-relative fallback.
  const baseFontPx = chartTextFontSizePx(def.fontSizeHpt, ptToPx)
    ?? Math.max(9, outerR * 0.09);

  const seriesBox = def.labelBox;

  // ── Build each label: wrapped lines + measured box + rim anchor ──────────
  const labels: PieCalloutLabel[] = [];
  let angle = startAngle;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const midAngle = angle + slice / 2;
    angle += slice;
    if (slice <= 0) continue;
    if (!indices.has(i)) continue;

    const ov = findOverride(i);
    // A genuine `<c:delete val="1"/>` (§21.2.2.43) skips the label; a per-point
    // *styling / flag* override is NOT a delete even though
    // it also has `text === ""` — key off the explicit `deleted` flag.
    if (dataLabelIsDeleted(def, ov)) continue;

    // §21.2.2.35 composition, with per-point `<c:dLbl>` show-flags (§21.2.2.47)
    // overriding the series defaults for this slice. Word stacks category name
    // and percent on separate lines, so each `show*` part is
    // its own line rather than space-joined.
    const showCatName = ov?.showCatName ?? def.showCatName;
    const showSerName = ov?.showSerName ?? def.showSerName;
    const showVal     = ov?.showVal ?? def.showVal;
    const showPercent = ov?.showPercent ?? def.showPercent;
    const showLegendKey = ov?.showLegendKey ?? def.showLegendKey ?? false;
    // Per-point overrides (font colour/size/bold + box), else series defaults.
    const fontPx = chartTextFontSizePx(ov?.fontSizeHpt, ptToPx) ?? baseFontPx;
    const bold = ov?.fontBold ?? def.fontBold ?? false;
    const labelFont = (ov?.fontFace ?? def.fontFace)
      ? chartFontFamily(chart, ov?.fontFace ?? def.fontFace, 'minor')
      : font;
    const fontColor = ov?.fontColor ? `#${ov.fontColor}` : (def.fontColor ? `#${def.fontColor}` : '#000');
    const box = mergeChartLabelBoxes(ov?.labelBox, seriesBox);
    const position = ov?.position ?? def.position ?? 'bestFit';

    const text = effectiveDataLabelText({
      customText: ov?.text,
      showCategory: showCatName,
      showSeries: showSerName,
      showValue: showVal,
      showPercent,
      category: (cats[i] ?? '').toString(),
      seriesName: s.name,
      sourceValue: vals[i],
      percentRatio: vals[i] / total,
      formatCode: ov?.formatCode ?? def.formatCode ?? s.valFormatCode ?? null,
      percentFormatCode: ov?.formatCode ?? def.formatCode ?? '0%',
      date1904: chart.date1904 ?? false,
      separator: ov?.separator ?? def.separator,
      defaultSeparator: '\n',
    });
    const legendKey = showLegendKey ? dataLabelLegendKey(0, i) : undefined;
    if (!text && !legendKey) continue;
    const textStyle = effectiveDataLabelTextStyle(ov, def);
    const richOptions = customRichDataLabelOptions(
      chart, ov, ptToPx, labelFont, bold, textStyle,
    );

    const authoredInsets = textStyle.textBodyAuthored === true
      || textStyle.textLInsEmu != null || textStyle.textTInsEmu != null
      || textStyle.textRInsEmu != null || textStyle.textBInsEmu != null;
    const bodyInsets = dataLabelInsets(textStyle, ptToPx);
    const padLeft = authoredInsets ? bodyInsets.left : Math.max(4, fontPx * 0.45);
    const padRight = authoredInsets ? bodyInsets.right : Math.max(4, fontPx * 0.45);
    const padTop = authoredInsets ? bodyInsets.top : Math.max(2, fontPx * 0.28);
    const padBottom = authoredInsets ? bodyInsets.bottom : Math.max(2, fontPx * 0.28);
    const lineGap = fontPx * 0.22;
    const lineH = fontPx + lineGap;
    ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontPx}px ${labelFont}`;
    const rich = richOptions
      ? resolveRichDataLabelBlock(ctx, richOptions, fontPx, fontColor)
      : null;
    if (richOptions && !rich) continue;
    let lines = rich ? [] : fitStyledDataLabelLines(
      text,
      Math.max(0, boundsW - padLeft - padRight),
      Math.max(0, boundsH - padTop - padBottom),
      lineH,
      value => ctx.measureText(value).width,
      textStyle,
    );
    if (!rich && lines.length === 0 && !legendKey) continue;
    let textW = rich?.width ?? 0;
    if (!rich) for (const ln of lines) textW = Math.max(textW, ctx.measureText(ln).width);
    const keyW = legendKey
      ? (legendSwatchWidths([legendKey.entry], fontPx, ptToPx)[0] ?? 0)
      : 0;
    const keyH = legendKey ? legendSwatchHeight(legendKey.entry, fontPx, ptToPx) : 0;
    const keyGap = legendKey && text ? LEGEND_SWATCH_TEXT_GAP : 0;
    let unrotatedW = keyW + keyGap + textW + padLeft + padRight;
    let unrotatedH = Math.max(
      keyH, rich?.height ?? (lines.length > 0 ? lines.length * lineH - lineGap : 0),
    ) + padTop + padBottom;
    let rotated = rotatedDataLabelSize(
      unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
    );
    let boxW = Math.min(rotated.w, boundsW);
    let boxH = Math.max(keyH, rich?.height ?? (lines.length > 0 ? lines.length * lineH - lineGap : 0));
    boxH = Math.min(rotated.h, boundsH);

    const rimX = cx2 + Math.cos(midAngle) * outerR;
    const rimY = cy2 + Math.sin(midAngle) * outerR;
    let leftSide = Math.cos(midAngle) < 0;

    // Initial box centre: outside the rim along the mid-angle. The gap scales
    // with the box so small slices get pulled further out (Word `bestFit`).
    const outGap = Math.max(boxW, boxH) * 0.55 + outerR * 0.06;
    let cxBox = rimX + Math.cos(midAngle) * outGap;
    let cyBox = rimY + Math.sin(midAngle) * outGap;
    let manualClip: DataLabelRect | undefined;
    let inside = false;
    if (ov?.manualLayout) {
      const manual = resolveDataLabelPlacement(
        { kind: 'point', x: cxBox, y: cyBox, position: 'ctr' },
        { x: boundsX, y: boundsY, w: boundsW, h: boundsH },
        { w: boxW, h: boxH },
        fontPx,
        ov.manualLayout,
        { x: chartX, y: chartY, w: chartW, h: chartH },
      );
      if (!manual) continue;
      boxW = manual.rect.w;
      boxH = manual.rect.h;
      unrotatedW = boxW;
      unrotatedH = boxH;
      if (!rich) {
        lines = fitStyledDataLabelLines(
          text,
          Math.max(0, boxW - padLeft - padRight - keyW - keyGap),
          Math.max(0, boxH - padTop - padBottom),
          lineH,
          value => ctx.measureText(value).width,
          textStyle,
        );
        if (lines.length === 0 && !legendKey) continue;
      }
      cxBox = manual.rect.x + manual.rect.w / 2;
      cyBox = manual.rect.y + manual.rect.h / 2;
      leftSide = cxBox < cx2;
      manualClip = manual.clip;
    } else if (position !== 'bestFit' && position !== 'outEnd') {
      const labelR = innerR > 0.01
        ? (innerR + outerR) / 2
        : outerR * PIE_CTR_LABEL_RADIUS_FRAC;
      const labelX = cx2 + Math.cos(midAngle) * labelR;
      const labelY = cy2 + Math.sin(midAngle) * labelR;
      const tangentialCapacity = 2 * labelR
        * Math.sin(Math.min(Math.PI, Math.abs(slice)) / 2) - fontPx;
      const radialCapacity = innerR > 0.01
        ? outerR - innerR - fontPx
        : outerR - fontPx;
      const sliceBounds = dataLabelRectIntersection(
        {
          x: labelX - tangentialCapacity / 2,
          y: labelY - radialCapacity / 2,
          w: tangentialCapacity,
          h: radialCapacity,
        },
        { x: boundsX, y: boundsY, w: boundsW, h: boundsH },
      );
      if (!sliceBounds) continue;
      if (!rich) {
        lines = fitStyledDataLabelLines(
          text,
          Math.max(0, sliceBounds.w - padLeft - padRight - keyW - keyGap),
          Math.max(0, sliceBounds.h - padTop - padBottom),
          lineH,
          value => ctx.measureText(value).width,
          textStyle,
        );
        if (lines.length === 0 && !legendKey) continue;
        textW = lines.reduce((width, line) => Math.max(width, ctx.measureText(line).width), 0);
        unrotatedW = keyW + keyGap + textW + padLeft + padRight;
        unrotatedH = Math.max(keyH, lines.length > 0 ? lines.length * lineH - lineGap : 0)
          + padTop + padBottom;
        rotated = rotatedDataLabelSize(
          unrotatedW, unrotatedH, textStyle.textRotation, textStyle.textVerticalMode,
        );
        boxW = rotated.w;
        boxH = rotated.h;
      } else {
        boxW = Math.min(boxW, sliceBounds.w);
        boxH = Math.min(boxH, sliceBounds.h);
      }
      const anchorPosition = position === 'inBase' || position === 'inEnd'
        ? 'ctr'
        : position;
      const placement = resolveDataLabelPlacement(
        { kind: 'point', x: labelX, y: labelY, position: anchorPosition },
        sliceBounds,
        { w: boxW, h: boxH },
        fontPx,
      );
      if (!placement) continue;
      cxBox = placement.textAlign === 'left'
        ? placement.x + boxW / 2
        : placement.textAlign === 'right'
          ? placement.x - boxW / 2
          : placement.x;
      cyBox = placement.textBaseline === 'top'
        ? placement.y + boxH / 2
        : placement.textBaseline === 'bottom'
          ? placement.y - boxH / 2
          : placement.y;
      leftSide = cxBox < cx2;
      manualClip = placement.clip;
      inside = true;
    }

    labels.push({
      lines, rich: rich ?? undefined, legendKey, lineHeight: lineH, midAngle, rimX, rimY,
      boxW, boxH, unrotatedW, unrotatedH, cxBox, cyBox,
      leftSide, fontColor, box, fontPx, bold, font: labelFont, textStyle, ptToPx,
      inside, manualClip,
    });
  }

  // ── Collision pass (bestFit): split into left/right columns and push boxes
  //    apart vertically so their rectangles do not overlap. Word lays labels
  //    out radially then de-overlaps; this greedy top-down separation +
  //    within-bounds fit-back is a faithful, deterministic approximation (no
  //    sample-specific tuning). ──
  const topLimit = boundsY + 2;
  const bottomLimit = boundsY + boundsH - 2;
  const band = bottomLimit - topLimit;
  const separate = (col: PieCalloutLabel[]): void => {
    if (col.length === 0) return;
    col.sort((a, b) => a.cyBox - b.cyBox);
    // Total height the boxes need when stacked edge-to-edge with a 3px gap
    // between them: the sum of box heights plus the inter-box gaps.
    let stackH = 0;
    for (const l of col) stackH += l.boxH;
    stackH += (col.length - 1) * 3;

    if (stackH > band) {
      // More label than plot: the boxes cannot all fit with the full 3px gaps
      // inside the plot rect. Distribute them so the FIRST box top sits at
      // topLimit and the LAST box bottom sits at bottomLimit, spacing the
      // in-between boxes by an equal step. This keeps the whole column WITHIN
      // [topLimit, bottomLimit] — never spilling past the bottom — which is the
      // overflow #767 guarded against. When the boxes are short enough to fit
      // (sumBoxH ≤ band) the step is a positive gap (no overlap); only a genuine
      // over-pack (sumBoxH > band, i.e. more labels than the plot can hold)
      // forces the boxes to touch/slightly overlap rather than escape the frame.
      const sumBoxH = col.reduce((a, l) => a + l.boxH, 0);
      const n = col.length;
      if (n === 1) {
        col[0].cyBox = Math.min(Math.max(col[0].cyBox, topLimit + col[0].boxH / 2), bottomLimit - col[0].boxH / 2);
        return;
      }
      // Equal gap so first-top = topLimit and last-bottom = bottomLimit:
      //   topLimit + ΣboxH + (n−1)·gap = bottomLimit  ⇒  gap = (band − ΣboxH)/(n−1)
      const gap = (band - sumBoxH) / (n - 1); // may be negative when over-packed
      let cursor = topLimit;
      for (const l of col) {
        l.cyBox = cursor + l.boxH / 2;
        cursor += l.boxH + gap;
      }
      return;
    }

    // Fits: push each box below the previous one by at least their combined half
    // heights (+ a small gap) so rectangles never overlap.
    for (let k = 1; k < col.length; k++) {
      const prev = col[k - 1];
      const cur = col[k];
      const minGap = (prev.boxH + cur.boxH) / 2 + 3;
      if (cur.cyBox - prev.cyBox < minGap) cur.cyBox = prev.cyBox + minGap;
    }
    // The overlap push above is one-directional (boxes only move DOWN), so a
    // bottom-heavy initial layout can now overrun EITHER bound. Because we are
    // in the fits case (stackH ≤ band) the rigid column is shorter than the
    // band, so a single slide brings BOTH ends inside [topLimit, bottomLimit] at
    // once. Slide up by any bottom overflow, then — symmetrically — down by any
    // top underflow. Sliding the whole column down cannot re-cross the bottom
    // because the column fits, so this two-step slide is a true round-trip
    // clamp (the earlier code capped the down-slide against a bottom "room" that
    // the prior up-slide had already zeroed, so a top underflow of ~100px was
    // left uncorrected — #767 was asymmetric, guarding only the bottom edge).
    const bottomOverflow = (col[col.length - 1].cyBox + col[col.length - 1].boxH / 2) - bottomLimit;
    if (bottomOverflow > 0) for (const l of col) l.cyBox -= bottomOverflow;
    const topUnderflow = topLimit - (col[0].cyBox - col[0].boxH / 2);
    if (topUnderflow > 0) for (const l of col) l.cyBox += topUnderflow;
  };
  separate(labels.filter(l => !l.manualClip && !l.leftSide));
  separate(labels.filter(l => !l.manualClip && l.leftSide));

  // Final round-trip clamp (both edges): guarantee no box escapes the plot rect
  // vertically, independent of which separate() branch ran. In the fits case the
  // symmetric slide above already lands every box inside [topLimit, bottomLimit];
  // in the over-packed case the equal-step distribution pins the first top to
  // topLimit and last bottom to bottomLimit. This per-box clamp is therefore a
  // no-op on the current paths, but makes the "no box leaves the frame at either
  // end" invariant explicit and robust to future layout changes. Clamp top FIRST
  // then bottom so a box taller than the band (degenerate) pins to the TOP edge
  // rather than escaping upward.
  for (const l of labels) {
    if (l.manualClip) continue;
    l.cyBox = Math.max(topLimit + l.boxH / 2, l.cyBox);
    l.cyBox = Math.min(bottomLimit - l.boxH / 2, l.cyBox);
  }

  // Horizontal clamp: keep each box fully inside the chart rect.
  const leftLimit = boundsX + 2;
  const rightLimit = boundsX + boundsW - 2;
  for (const l of labels) {
    if (l.manualClip) continue;
    const half = l.boxW / 2;
    if (l.cxBox - half < leftLimit) l.cxBox = leftLimit + half;
    if (l.cxBox + half > rightLimit) l.cxBox = rightLimit - half;
  }

  // ── Draw leader lines first (under the boxes), then boxes + text ─────────
  ctx.save();
  ctx.beginPath();
  ctx.rect(boundsX, boundsY, boundsW, boundsH);
  ctx.clip();
  const leader = chartStyleRoleLeaderLine(chart, def);
  const leaderColor = leader.color ? `#${leader.color}` : '#a6a6a6';
  const leaderPx = leader.widthEmu
    ? Math.max(0.5, (leader.widthEmu / EMU_PER_PT) * ptToPx)
    : 1;
  ctx.setLineDash(dashPatternForPreset(leader.dash ?? undefined, leaderPx));

  for (const l of labels) {
    // The box edge nearest the pie centre — where a leader line should meet.
    const edgeX = l.cxBox + (l.leftSide ? l.boxW / 2 : -l.boxW / 2);
    const edgeY = l.cyBox;
    // Distance from the box's inner edge to its slice rim. When the box abuts
    // the slice the leader is redundant; draw one only past a small threshold.
    const dx = edgeX - l.rimX;
    const dy = edgeY - l.rimY;
    const dist = Math.hypot(dx, dy);
    if (!l.inside && def.showLeaderLines && leader.hidden !== true
      && (leader.paintAuthored !== true || leader.color != null)
      && dist > l.fontPx * 0.9) {
      ctx.beginPath();
      ctx.moveTo(l.rimX, l.rimY);
      ctx.lineTo(edgeX, edgeY);
      ctx.strokeStyle = leaderColor;
      ctx.lineWidth = leaderPx;
      ctx.stroke();
    }
  }

  for (const l of labels) {
    if (l.manualClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(l.manualClip.x, l.manualClip.y, l.manualClip.w, l.manualClip.h);
      ctx.clip();
    }
    const bx = l.cxBox - l.boxW / 2;
    const by = l.cyBox - l.boxH / 2;
    // Box fill + border (§21.2.2.197 spPr). Fill may carry an 8-digit RGBA hex
    // (e.g. a 90%-opacity white) — valid canvas fillStyle.
    paintChartLabelBox(
      ctx, l.box, { x: bx, y: by, w: l.boxW, h: l.boxH }, ptToPx,
      shapeRotationDeg,
    );
    const authoredInsets = l.textStyle.textBodyAuthored === true
      || l.textStyle.textLInsEmu != null || l.textStyle.textTInsEmu != null
      || l.textStyle.textRInsEmu != null || l.textStyle.textBInsEmu != null;
    const bodyInsets = dataLabelInsets(l.textStyle, l.ptToPx);
    const padLeft = authoredInsets ? bodyInsets.left : Math.max(4, l.fontPx * 0.45);
    const padRight = authoredInsets ? bodyInsets.right : Math.max(4, l.fontPx * 0.45);
    const padTop = authoredInsets ? bodyInsets.top : Math.max(2, l.fontPx * 0.28);
    const padBottom = authoredInsets ? bodyInsets.bottom : Math.max(2, l.fontPx * 0.28);
    const rotated = rotatedDataLabelSize(
      l.unrotatedW, l.unrotatedH,
      l.textStyle.textRotation, l.textStyle.textVerticalMode,
    );
    const contentCx = l.cxBox + (padLeft - padRight) / 2;
    const contentCy = l.cyBox + (padTop - padBottom) / 2;
    const innerLeft = bx + padLeft;
    const innerRight = bx + l.boxW - padRight;
    const innerWidth = Math.max(0, innerRight - innerLeft);
    const paintAlign = dataLabelCanvasTextAlign(l.textStyle, 'center');
    const alignedX = paintAlign === 'left' ? innerLeft
      : paintAlign === 'right' ? innerRight : contentCx;
    const anchoredCenterY = (contentHeight: number): number =>
      (l.textStyle.textVerticalAnchor
        ?? (l.textStyle.textBodyAuthored === true ? 't' : 'ctr')) === 't'
        ? by + padTop + contentHeight / 2
        : (l.textStyle.textVerticalAnchor
          ?? (l.textStyle.textBodyAuthored === true ? 't' : 'ctr')) === 'b'
          ? by + l.boxH - padBottom - contentHeight / 2
          : contentCy;
    const alignedGroupLeft = (contentWidth: number): number =>
      paintAlign === 'left' ? innerLeft
        : paintAlign === 'right' ? innerRight - contentWidth
          : contentCx - contentWidth / 2;
    ctx.save();
    ctx.beginPath();
    ctx.rect(bx, by, l.boxW, l.boxH);
    ctx.clip();
    if (rotated.radians !== 0) {
      ctx.translate(l.cxBox, l.cyBox);
      ctx.rotate(rotated.radians);
      ctx.translate(-l.cxBox, -l.cyBox);
    }
    // Text: centred, stacked lines. A custom rich body uses the same bounded
    // inline block that measured the box, keeping measurement and paint exact.
    if (!l.legendKey) {
      const textHeight = l.rich?.height
        ?? Math.max(0, l.lines.length * l.lineHeight - (l.lineHeight - l.fontPx));
      const textCenterY = anchoredCenterY(textHeight);
      if (l.rich) {
        paintRichDataLabelBlock(
          ctx, l.rich, alignedX, textCenterY, paintAlign, 'middle', innerWidth,
        );
        ctx.restore();
        if (l.manualClip) ctx.restore();
        continue;
      }
      ctx.font = `${l.textStyle.fontItalic ? 'italic ' : ''}${l.bold ? 'bold ' : ''}${l.fontPx}px ${l.font}`;
      ctx.fillStyle = l.fontColor;
      ctx.textAlign = paintAlign;
      ctx.textBaseline = 'middle';
      const lineGap = l.lineHeight - l.fontPx;
      const baselineShift = (l.textStyle.fontBaseline ?? 0) * l.fontPx;
      const blockTop = textCenterY
        - (l.lines.length * l.lineHeight - lineGap) / 2 + l.fontPx / 2 - baselineShift;
      if (!(l.textStyle.fontPaintAuthored === true
        && (l.textStyle.fontHidden === true || l.textStyle.fontColor == null))) {
        for (let li = 0; li < l.lines.length; li++) {
          ctx.fillText(l.lines[li], alignedX, blockTop + li * l.lineHeight);
        }
      }
      ctx.restore();
      if (l.manualClip) ctx.restore();
      continue;
    }
    const keyWidth = l.legendKey
      ? (legendSwatchWidths([l.legendKey.entry], l.fontPx, l.legendKey.ptToPx)[0] ?? 0)
      : 0;
    const keyHeight = l.legendKey
      ? legendSwatchHeight(l.legendKey.entry, l.fontPx, l.legendKey.ptToPx)
      : 0;
    const keyGap = l.legendKey && (l.rich || l.lines.length > 0) ? LEGEND_SWATCH_TEXT_GAP : 0;
    const textWidth = l.rich?.width ?? l.lines.reduce(
      (width, line) => Math.max(width, ctx.measureText(line).width), 0,
    );
    const groupWidth = keyWidth + keyGap + textWidth;
    const groupHeight = Math.max(
      keyHeight,
      l.rich?.height ?? Math.max(0, l.lines.length * l.lineHeight - (l.lineHeight - l.fontPx)),
    );
    const groupCenterY = anchoredCenterY(groupHeight);
    const contentLeft = alignedGroupLeft(groupWidth);
    if (l.legendKey) {
      drawLegendSwatch(
        ctx,
        l.legendKey.entry.swatchStyle,
        l.legendKey.entry.color,
        contentLeft,
        groupCenterY - keyHeight / 2,
        keyWidth,
        keyHeight,
        l.legendKey.entry.marker,
        l.legendKey.entry.fillPaint,
        l.legendKey.entry.outlinePaint,
        l.legendKey.entry.outlineColor,
        l.legendKey.entry.outlineWidthEmu,
        l.legendKey.entry.outlineDash,
        l.legendKey.entry.outlineCustomDash,
        l.legendKey.entry.outlineCap,
        l.legendKey.entry.outlineJoin,
        l.legendKey.ptToPx,
        l.legendKey.shapeRotationDeg,
        l.legendKey.entry.directEffect,
        l.legendKey.entry.fallbackEffect,
        l.legendKey.entry.directEffectIndex,
        l.legendKey.entry.fallbackEffectIndex,
      );
    }
    if (l.rich) {
      paintRichDataLabelBlock(
        ctx, l.rich, contentLeft + keyWidth + keyGap, groupCenterY, 'left', 'middle',
        textWidth,
      );
      ctx.restore();
      if (l.manualClip) ctx.restore();
      continue;
    }
    ctx.font = `${l.textStyle.fontItalic ? 'italic ' : ''}${l.bold ? 'bold ' : ''}${l.fontPx}px ${l.font}`;
    ctx.fillStyle = l.fontColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    const lineGap = l.lineHeight - l.fontPx;
    const baselineShift = (l.textStyle.fontBaseline ?? 0) * l.fontPx;
    const blockTop = groupCenterY
      - (l.lines.length * l.lineHeight - lineGap) / 2 + l.fontPx / 2 - baselineShift;
    if (!(l.textStyle.fontPaintAuthored === true
      && (l.textStyle.fontHidden === true || l.textStyle.fontColor == null))) {
      for (let li = 0; li < l.lines.length; li++) {
        ctx.fillText(l.lines[li], contentLeft + keyWidth + keyGap, blockTop + li * l.lineHeight);
      }
    }
    ctx.restore();
    if (l.manualClip) ctx.restore();
  }
  ctx.restore();
}
