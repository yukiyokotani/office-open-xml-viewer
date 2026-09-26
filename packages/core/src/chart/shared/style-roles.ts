// Classic chart style roles helpers.
import type { ChartDataLabelOverride, ChartDecorationLineStyle, ChartExElementStyle, ChartLabelBox, ChartModel, ChartRect, ChartSeries, ChartSeriesDataLabels, ChartStockUpDownBarStyle, ChartStyleRole, ChartTrendline, SecondaryValueAxis } from '../../types/chart';
import { resolveFill } from '../../shape/paint.js';
import { axisLineWidthPx } from '../axis-style.js';
import { chartSeriesVariesByPoint, chartStyleDashChoice, effectiveChartStyleRole, rawLinkedChartStyleRole } from '../effective-style.js';
import { chartStyleDirectFillDecision, chartStyleDirectLineDecision, chartStyleDirectNoFillDecision, chartStyleDirectNoLineDecision, chartStyleFillCascade, chartStyleFontColor, chartStyleLineCascade, chartStyleLineDecision } from '../style-paint.js';
import { chartStyleEffectOwner } from '../style-effects.js';
import { hasVisiblePointMarkerOverride, pointHasMarkerDetail } from '../marker-style.js';
import { chartLabelBoxHasVisiblePaint, effectiveChartLabelBoxFill, mergeChartLabelBoxes } from '../label-box.js';
import { EMU_PER_PT } from '../../units.js';
import { indexChartPlotGroups } from '../plot-groups.js';
import { dashPatternForPreset } from './geometry.js';
import type { ChartExStyle } from './chartex-style.js';
import { chartExSeriesFormatIndex } from './palette.js';


export function applyDecorationLineStyle(
  ctx: CanvasRenderingContext2D,
  style: ChartDecorationLineStyle,
  ptToPx: number,
  bounds: ChartRect = {
    x: 0,
    y: 0,
    w: Math.max(1, ctx.canvas?.width ?? 1),
    h: Math.max(1, ctx.canvas?.height ?? 1),
  },
  shapeRotationDeg = 0,
): boolean {
  if (style.hidden === true
    || (style.paintAuthored === true && style.color == null && style.fill == null)) return false;
  const stroke = style.fill != null
    ? resolveFill(
        style.fill, ctx, bounds.x, bounds.y, bounds.w, bounds.h, shapeRotationDeg,
      )
    : style.color != null ? `#${style.color}` : '#000000';
  if (stroke == null) return false;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = style.widthEmu != null
    ? axisLineWidthPx(style.widthEmu, ptToPx)
    : Math.max(1, 0.75 * ptToPx);
  ctx.setLineDash(dashPatternForPreset(style.dash ?? undefined, ctx.lineWidth));
  ctx.lineCap = style.cap === 'rnd' ? 'round' : style.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = style.join === 'round' || style.join === 'bevel' ? style.join : 'miter';
  return true;
}


export function chartStyleRoleLine(
  chart: ChartModel,
  direct: ChartDecorationLineStyle,
  role: ChartStyleRole,
  compatibility?: ChartExElementStyle,
): ChartDecorationLineStyle {
  // A bounded Office compatibility recipe is a host/family default, not
  // authored OOXML. It therefore sits between the raw linked role and the
  // ECMA numeric role. `chartStyleRoles` is already linked-over-numeric and
  // cannot express that extra layer without recovering the retained sources.
  const linked = compatibility
    ? effectiveChartStyleRole(
        effectiveChartStyleRole(chart.classicChartStyleRoles?.[role], compatibility),
        chart.linkedChartStyleRoles?.[role] ?? (
          chart.classicChartStyleRoles == null ? chart.chartStyleRoles?.[role] : undefined
        ),
      )
    : chart.chartStyleRoles?.[role];
  const rawLinked = rawLinkedChartStyleRole(chart, role);
  const directPaintAuthored = direct.paintAuthored === true
    || direct.fill != null || direct.color != null || direct.hidden === true;
  const directStyleLine = chartStyleDirectLineDecision(direct.style, rawLinked, 0);
  const directNoLine = direct.hidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const lineDecision = directNoLine !== undefined ? directNoLine
    : direct.fill ?? (direct.color ? { fillType: 'solid' as const, color: direct.color }
      : directStyleLine !== undefined
        ? directStyleLine
        : direct.paintAuthored === true && direct.hidden !== true
        ? null
        : chartStyleLineCascade(linked, rawLinked, 0, direct.style));
  const dash = chartStyleDashChoice(
    direct.dash != null ? { lineDash: direct.dash, lineDashAuthored: true } : undefined,
    direct.style,
    linked,
  );
  return {
    style: chartStyleEffectOwner(direct.style, linked),
    fill: lineDecision != null && lineDecision.fillType !== 'solid'
      ? lineDecision : null,
    color: lineDecision?.fillType === 'solid' ? lineDecision.color : null,
    paintAuthored: directPaintAuthored
      ? direct.paintAuthored
      : lineDecision !== undefined ? true : undefined,
    widthEmu: direct.widthEmu ?? direct.style?.lineWidthEmu ?? linked?.lineWidthEmu ?? null,
    dash: dash?.lineDash ?? null,
    cap: direct.cap ?? direct.style?.lineCap ?? linked?.lineCap ?? null,
    join: direct.join ?? direct.style?.lineJoin ?? linked?.lineJoin ?? null,
    hidden: lineDecision === null ? true : null,
  };
}


export function chartStyleRoleBarPaint(
  chart: ChartModel,
  direct: ChartStockUpDownBarStyle['up'],
  role: 'upBar' | 'downBar',
  automaticPaint?: {
    lineColor: string;
    lineWidthEmu: number;
    upFillColor: string;
    downFillColor: string;
  },
): ChartStockUpDownBarStyle['up'] {
  const compatibility: ChartExElementStyle | undefined = automaticPaint ? {
    fillColors: [role === 'upBar' ? automaticPaint.upFillColor : automaticPaint.downFillColor],
    fillPaintAuthored: true,
    lineColors: [automaticPaint.lineColor],
    linePaintAuthored: true,
    lineWidthEmu: automaticPaint.lineWidthEmu,
  } : undefined;
  const linked = compatibility
    ? effectiveChartStyleRole(
        effectiveChartStyleRole(chart.classicChartStyleRoles?.[role], compatibility),
        chart.linkedChartStyleRoles?.[role] ?? (
          chart.classicChartStyleRoles == null ? chart.chartStyleRoles?.[role] : undefined
        ),
      )
    : chart.chartStyleRoles?.[role];
  const rawLinked = rawLinkedChartStyleRole(chart, role);
  const directFillAuthored = direct.fillPaintAuthored === true
    || direct.fillColor != null || direct.fill != null || direct.fillHidden === true;
  const directLineAuthored = direct.linePaintAuthored === true
    || direct.lineColor != null || direct.lineHidden === true;
  const directStyleFill = chartStyleDirectFillDecision(direct.style, rawLinked, 0);
  const directStyleLine = chartStyleDirectLineDecision(direct.style, rawLinked, 0);
  const directNoFill = direct.fillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directNoLine = direct.lineHidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const fillDecision = directNoFill !== undefined ? directNoFill
    : direct.fill != null ? direct.fill
    : direct.fillColor != null ? { fillType: 'solid' as const, color: direct.fillColor }
    : directStyleFill !== undefined
      ? directStyleFill
      : direct.fillPaintAuthored === true && direct.fillHidden !== true ? null
      : chartStyleFillCascade(linked, rawLinked, 0, direct.style);
  const lineDecision = directNoLine !== undefined ? directNoLine
    : direct.lineColor ? { fillType: 'solid' as const, color: direct.lineColor }
      : directStyleLine !== undefined
        ? directStyleLine
        : direct.linePaintAuthored === true && direct.lineHidden !== true
        ? null
        : chartStyleLineCascade(linked, rawLinked, 0, direct.style);
  return {
    style: chartStyleEffectOwner(direct.style, linked),
    fillColor: fillDecision?.fillType === 'solid' ? fillDecision.color : null,
    fill: fillDecision != null && fillDecision.fillType !== 'solid'
      && fillDecision.fillType !== 'none' ? fillDecision : null,
    fillPaintAuthored: directFillAuthored
      ? direct.fillPaintAuthored
      : fillDecision !== undefined ? true : undefined,
    fillHidden: fillDecision === null ? true : null,
    lineColor: lineDecision?.fillType === 'solid' ? lineDecision.color : null,
    linePaintAuthored: directLineAuthored
      ? direct.linePaintAuthored
      : lineDecision !== undefined ? true : undefined,
    lineWidthEmu: direct.lineWidthEmu ?? linked?.lineWidthEmu ?? null,
    lineDash: direct.lineDash ?? linked?.lineDash ?? null,
    lineCap: direct.lineCap ?? linked?.lineCap ?? null,
    lineJoin: direct.lineJoin ?? linked?.lineJoin ?? null,
    lineHidden: lineDecision === null ? true : null,
  };
}


/** Draws the one-per-category drop-line envelope shared by classic line,
 * area, and stock charts. ECMA-376 assigns the geometry to the owning chart
 * group; each envelope joins its effective category-axis crossing to every
 * finite plotted point at that category. */
export function drawDropLineEnvelopes(
  ctx: CanvasRenderingContext2D,
  members: ChartSeries[],
  pointCount: number,
  toX: (index: number) => number,
  yMapFor: (series: ChartSeries) => (value: number) => number,
  categoryAxisYFor: (series: ChartSeries) => number,
  valueFor: (series: ChartSeries, index: number) => number | null,
): void {
  for (let index = 0; index < pointCount; index++) {
    let minY = Infinity;
    let maxY = -Infinity;
    let hasPoint = false;
    for (const series of members) {
      const value = valueFor(series, index);
      if (value == null || !Number.isFinite(value)) continue;
      const pointY = yMapFor(series)(value);
      const axisY = categoryAxisYFor(series);
      if (!Number.isFinite(pointY) || !Number.isFinite(axisY)) continue;
      minY = Math.min(minY, pointY, axisY);
      maxY = Math.max(maxY, pointY, axisY);
      hasPoint = true;
    }
    if (!hasPoint || Math.abs(maxY - minY) < 0.01) continue;
    ctx.beginPath();
    ctx.moveTo(toX(index), minY);
    ctx.lineTo(toX(index), maxY);
    ctx.stroke();
  }
}


export function chartStyleRoleErrorBar(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['errBars']>[number],
): NonNullable<ChartSeries['errBars']>[number] {
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.color,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.dash,
    hidden: direct.hidden,
  }, 'errorBar');
  return {
    ...direct,
    color: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    dash: linked.dash ?? undefined,
    hidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}


export function chartStyleRoleLeaderLine(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    style: direct.leaderLineStyle,
    color: direct.leaderLineColor,
    paintAuthored: direct.leaderLinePaintAuthored,
    widthEmu: direct.leaderLineWidthEmu,
    dash: direct.leaderLineDash,
    hidden: direct.leaderLineHidden,
  }, 'leaderLine');
}


export function chartStyleRoleTrendline(
  chart: ChartModel,
  direct: NonNullable<ChartSeries['trendLines']>[number],
): NonNullable<ChartSeries['trendLines']>[number] {
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.lineColor,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.lineDash,
    hidden: direct.lineHidden,
  }, 'trendline');
  return {
    ...direct,
    lineColor: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    lineDash: linked.dash ?? undefined,
    lineHidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}


export function chartStyleRoleDataTable(
  chart: ChartModel,
  direct: NonNullable<ChartModel['dataTable']>,
): NonNullable<ChartModel['dataTable']> {
  const role = chart.chartStyleRoles?.dataTable;
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataTable');
  const linked = chartStyleRoleLine(chart, {
    style: direct.style,
    color: direct.lineColor,
    paintAuthored: direct.linePaintAuthored,
    widthEmu: direct.lineWidthEmu,
    dash: direct.lineDash,
    hidden: direct.lineHidden,
  }, 'dataTable');
  const directStyleFill = chartStyleDirectFillDecision(direct.style, rawLinked, 0);
  const directNoFill = direct.fillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const fillDecision = directNoFill !== undefined ? directNoFill
    : direct.fill ?? (direct.fillColor ? { fillType: 'solid' as const, color: direct.fillColor }
      : directStyleFill !== undefined
        ? directStyleFill
        : direct.fillPaintAuthored === true && direct.fillHidden !== true
        ? null
        : chartStyleFillCascade(role, rawLinked, 0, direct.style));
  const fill = fillDecision != null && fillDecision.fillType !== 'solid'
    && fillDecision.fillType !== 'image' && fillDecision.fillType !== 'none'
    ? fillDecision : null;
  const fillColor = fillDecision?.fillType === 'solid' ? fillDecision.color : null;
  const fillHidden = fillDecision === null ? true : null;
  const fillPaintAuthored = direct.fillPaintAuthored
    ?? (fillDecision !== undefined ? true : undefined);
  const textPaint = effectiveInheritedChartTextPaint(chart,
    direct.fontColor,
    direct.fontPaintAuthored,
    role,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? role?.fontSizeHpt,
    fontBold: direct.fontBold ?? chart.chartTextStyle?.fontBold ?? role?.fontBold,
    fontItalic: direct.fontItalic ?? chart.chartTextStyle?.fontItalic ?? role?.fontItalic,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: direct.fontPaintAuthored === true
      ? direct.fontHidden
      : role?.fontHidden,
    fontFace: direct.fontFace ?? chart.chartTextStyle?.fontFace ?? role?.fontFace,
    fill,
    fillColor,
    fillHidden,
    fillPaintAuthored,
    lineColor: linked.color ?? undefined,
    lineWidthEmu: linked.widthEmu ?? undefined,
    lineDash: linked.dash ?? undefined,
    lineHidden: linked.hidden ?? undefined,
    linePaintAuthored: linked.paintAuthored,
  };
}


export interface LinkedGridlineResult {
  visible: boolean | null | undefined;
  color?: string | null;
  widthEmu?: number | null;
  dash?: string | null;
  paintAuthored?: boolean | null;
}


export function chartStyleRoleGridline(
  chart: ChartModel,
  role: 'gridlineMajor' | 'gridlineMinor',
  visible: boolean | null | undefined,
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  paintAuthored: boolean | null | undefined,
  directStyle?: ChartExStyle | null,
): LinkedGridlineResult {
  if (visible !== true) {
    return { visible, color, widthEmu, dash, paintAuthored };
  }
  const linked = chartStyleRoleLine(
    chart, { style: directStyle, color, widthEmu, dash, paintAuthored }, role,
  );
  return {
    visible: linked.hidden !== true
      && !(linked.paintAuthored === true && linked.color == null),
    color: linked.color,
    widthEmu: linked.widthEmu,
    dash: linked.dash,
    paintAuthored: linked.paintAuthored,
  };
}


export function chartStyleRoleSecondaryGridlines(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
): SecondaryValueAxis | null | undefined {
  if (!axis) return axis;
  if (!chart.chartStyleRoles?.gridlineMajor && !chart.chartStyleRoles?.gridlineMinor) return axis;
  const major = chartStyleRoleGridline(
    chart, 'gridlineMajor', axis.majorGridlines,
    axis.majorGridlineColor, axis.majorGridlineWidthEmu, axis.majorGridlineDash,
    axis.majorGridlinePaintAuthored,
    axis.majorGridlineStyle,
  );
  const minor = chartStyleRoleGridline(
    chart, 'gridlineMinor', axis.minorGridlines,
    axis.minorGridlineColor, axis.minorGridlineWidthEmu, axis.minorGridlineDash,
    axis.minorGridlinePaintAuthored,
    axis.minorGridlineStyle,
  );
  const changed = major.visible !== axis.majorGridlines
    || major.color !== axis.majorGridlineColor
    || major.widthEmu !== axis.majorGridlineWidthEmu
    || major.dash !== axis.majorGridlineDash
    || major.paintAuthored !== axis.majorGridlinePaintAuthored
    || minor.visible !== axis.minorGridlines
    || minor.color !== axis.minorGridlineColor
    || minor.widthEmu !== axis.minorGridlineWidthEmu
    || minor.dash !== axis.minorGridlineDash
    || minor.paintAuthored !== axis.minorGridlinePaintAuthored;
  return changed ? {
    ...axis,
    majorGridlines: major.visible ?? undefined,
    majorGridlineColor: major.color,
    majorGridlineWidthEmu: major.widthEmu,
    majorGridlineDash: major.dash,
    majorGridlinePaintAuthored: major.paintAuthored,
    minorGridlines: minor.visible ?? undefined,
    minorGridlineColor: minor.color,
    minorGridlineWidthEmu: minor.widthEmu,
    minorGridlineDash: minor.dash,
    minorGridlinePaintAuthored: minor.paintAuthored,
  } : axis;
}


export function chartStyleRoleAxisLine(
  chart: ChartModel,
  role: 'categoryAxis' | 'valueAxis',
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  hidden: boolean,
  paintAuthored: boolean | null | undefined,
  directStyle?: ChartExStyle | null,
): ChartDecorationLineStyle {
  return chartStyleRoleLine(chart, {
    style: directStyle,
    color,
    widthEmu,
    dash,
    paintAuthored,
    // The shared axis model stores the effective boolean, so false means no
    // direct noFill rather than an authored visible override.
    hidden: hidden ? true : undefined,
  }, role);
}


/** The flat axis model can carry a resolved solid color but not an arbitrary
 * DrawingML stroke paint. Preserve authored ownership by suppressing an
 * unresolved/structured paint here instead of reviving the semantic black
 * fallback in `resolveAxisLine`. */
export function chartAxisLineIsHidden(line: ChartDecorationLineStyle): boolean {
  return line.hidden === true || (line.paintAuthored === true && line.color == null);
}


export function chartStyleRoleSecondaryAxisLine(
  chart: ChartModel,
  axis: SecondaryValueAxis | null | undefined,
  role: 'categoryAxis' | 'valueAxis',
): SecondaryValueAxis | null | undefined {
  const style = chart.chartStyleRoles?.[role];
  if (!axis || (!style && !chart.chartTextStyle)) return axis;
  const line = chartStyleRoleAxisLine(
    chart, role, axis.lineColor, axis.lineWidthEmu, axis.lineDash, axis.lineHidden,
    axis.linePaintAuthored, axis.style,
  );
  const lineHidden = chartAxisLineIsHidden(line);
  const fontSizeHpt = axis.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? style?.fontSizeHpt;
  const fontBold = axis.fontBold ?? chart.chartTextStyle?.fontBold ?? style?.fontBold;
  const fontItalic = axis.fontItalic ?? chart.chartTextStyle?.fontItalic ?? style?.fontItalic;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, axis.fontColor, axis.fontPaintAuthored, style,
  );
  const fontColor = textPaint.color;
  const fontFace = axis.fontFace ?? chart.chartTextStyle?.fontFace ?? style?.fontFace;
  const titleStyle = chart.chartStyleRoles?.axisTitle;
  const titleFontSizeHpt = axis.titleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? titleStyle?.fontSizeHpt;
  const titleFontBold = axis.titleFontBold ?? chart.chartTextStyle?.fontBold ?? titleStyle?.fontBold;
  const titleFontItalic = axis.titleFontItalic
    ?? chart.chartTextStyle?.fontItalic ?? titleStyle?.fontItalic;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart, axis.titleFontColor, axis.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color;
  const titleFontFace = axis.titleFontFace
    ?? chart.chartTextStyle?.fontFace ?? titleStyle?.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && textPaint.authored === axis.fontPaintAuthored
    && fontFace === axis.fontFace
    && titleFontSizeHpt === axis.titleFontSizeHpt
    && titleFontBold === axis.titleFontBold
    && titleFontItalic === axis.titleFontItalic
    && titleFontColor === axis.titleFontColor
    && titleTextPaint.authored === axis.titleFontPaintAuthored
    && titleFontFace === axis.titleFontFace) return axis;
  return {
    ...axis,
    lineColor: line.color,
    lineWidthEmu: line.widthEmu,
    lineDash: line.dash,
    linePaintAuthored: line.paintAuthored,
    lineHidden,
    fontSizeHpt,
    fontBold,
    fontItalic,
    fontColor,
    fontPaintAuthored: textPaint.authored,
    fontFace,
    titleFontSizeHpt,
    titleFontBold,
    titleFontItalic,
    titleFontColor,
    titleFontPaintAuthored: titleTextPaint.authored,
    titleFontFace,
  };
}


export function chartStyleRoleSeriesAxis(chart: ChartModel): ChartModel {
  const axis = chart.threeD?.seriesAxis;
  const style = chart.chartStyleRoles?.seriesAxis;
  if (!axis || (!style && !chart.chartTextStyle)) return chart;
  const line = chartStyleRoleLine(chart, {
    style: axis.style,
    color: axis.lineColor,
    widthEmu: axis.lineWidthEmu,
    dash: axis.lineDash,
    hidden: axis.lineHidden ? true : undefined,
    paintAuthored: axis.linePaintAuthored,
  }, 'seriesAxis');
  const lineHidden = chartAxisLineIsHidden(line);
  const fontSizeHpt = axis.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt ?? style?.fontSizeHpt;
  const fontBold = axis.fontBold ?? chart.chartTextStyle?.fontBold ?? style?.fontBold;
  const fontItalic = axis.fontItalic ?? chart.chartTextStyle?.fontItalic ?? style?.fontItalic;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, axis.fontColor, axis.fontPaintAuthored, style,
  );
  const fontColor = textPaint.color;
  const fontFace = axis.fontFace ?? chart.chartTextStyle?.fontFace ?? style?.fontFace;
  const titleStyle = chart.chartStyleRoles?.axisTitle;
  const titleFontSizeHpt = axis.titleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? titleStyle?.fontSizeHpt;
  const titleFontBold = axis.titleFontBold ?? chart.chartTextStyle?.fontBold ?? titleStyle?.fontBold;
  const titleFontItalic = axis.titleFontItalic
    ?? chart.chartTextStyle?.fontItalic ?? titleStyle?.fontItalic;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart, axis.titleFontColor, axis.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color;
  const titleFontFace = axis.titleFontFace
    ?? chart.chartTextStyle?.fontFace ?? titleStyle?.fontFace;
  if (line.color === axis.lineColor
    && line.widthEmu === axis.lineWidthEmu
    && line.dash === axis.lineDash
    && lineHidden === axis.lineHidden
    && fontSizeHpt === axis.fontSizeHpt
    && fontBold === axis.fontBold
    && fontItalic === axis.fontItalic
    && fontColor === axis.fontColor
    && textPaint.authored === axis.fontPaintAuthored
    && fontFace === axis.fontFace
    && titleFontSizeHpt === axis.titleFontSizeHpt
    && titleFontBold === axis.titleFontBold
    && titleFontItalic === axis.titleFontItalic
    && titleFontColor === axis.titleFontColor
    && titleTextPaint.authored === axis.titleFontPaintAuthored
    && titleFontFace === axis.titleFontFace) return chart;
  return {
    ...chart,
    threeD: {
      ...chart.threeD,
      seriesAxis: {
        ...axis,
        lineColor: line.color,
        lineWidthEmu: line.widthEmu,
        lineDash: line.dash,
        lineHidden,
        fontSizeHpt,
        fontBold,
        fontItalic,
        fontColor,
        fontPaintAuthored: textPaint.authored,
        fontFace,
        titleFontSizeHpt,
        titleFontBold,
        titleFontItalic,
        titleFontColor,
        titleFontPaintAuthored: titleTextPaint.authored,
        titleFontFace,
      },
    },
  };
}


export function isClassicMarkerSeries(
  chart: ChartModel,
  series: ChartSeries,
  group: NonNullable<ChartModel['plotGroups']>[number] | undefined,
): boolean {
  if (group?.kind === 'bubble' || (group == null && chart.chartType === 'bubble')) return false;
  const family = group?.kind === 'scatter' ? 'scatter' : series.seriesType ?? chart.chartType;
  return family === 'line'
    || family === 'stackedLine'
    || family === 'stackedLinePct'
    || family === 'area'
    || family === 'stackedArea'
    || family === 'stackedAreaPct'
    || family === 'scatter'
    || family === 'radar'
    || family === 'stock';
}


export function chartStyleRoleMarker(
  chart: ChartModel,
  direct: ChartSeries,
  index: number,
  count: number,
  group: NonNullable<ChartModel['plotGroups']>[number] | undefined,
): ChartSeries {
  void count;
  // A point-varying group uses a separate formatting-index domain. Do not
  // flatten that role into series defaults here: drawChartMarker selects the
  // point role at paint time so every marker keeps its own index.
  const linked = chartSeriesVariesByPoint(chart, index)
    ? undefined
    : chart.chartStyleRoles?.dataPointMarker;
  const rawLinked = chartSeriesVariesByPoint(chart, index)
    ? undefined
    : rawLinkedChartStyleRole(chart, 'dataPointMarker');
  if (!isClassicMarkerSeries(chart, direct, group)
    || ((direct.showMarker === false || direct.markerSymbol === 'none')
      && !hasVisiblePointMarkerOverride(direct))) return direct;
  const styleIndex = chartExSeriesFormatIndex(direct, index);
  const directStyleFill = chartStyleDirectFillDecision(
    direct.markerStyle, rawLinked, styleIndex,
  );
  const directFillAuthored = direct.markerFillPaintAuthored === true
      && direct.markerStyle?.fillHidden !== true
    || direct.markerFill != null || direct.markerFillPaint !== undefined
    || directStyleFill !== undefined;
  const effectiveFill = directFillAuthored
    ? directStyleFill
    : chartStyleFillCascade(linked, rawLinked, styleIndex, direct.markerStyle);
  const markerFill = direct.markerFill
    ?? (effectiveFill?.fillType === 'solid' ? effectiveFill.color
      : effectiveFill === null ? '00000000' : null);
  const markerFillPaint = direct.markerFillPaint !== undefined
    ? direct.markerFillPaint
    : effectiveFill?.fillType === 'gradient'
        || effectiveFill?.fillType === 'pattern'
        || effectiveFill?.fillType === 'image'
      ? effectiveFill
      : undefined;
  const markerFillPaintAuthored = directFillAuthored
    ? direct.markerFillPaintAuthored
    : effectiveFill !== undefined
      ? true
      : undefined;
  const directStyleLine = chartStyleDirectLineDecision(
    direct.markerStyle, rawLinked, styleIndex,
  );
  const directLineAuthored = direct.markerLine != null
    || direct.markerLinePaintAuthored === true && direct.markerStyle?.lineHidden !== true
    || directStyleLine !== undefined;
  const effectiveLine = directLineAuthored
    ? directStyleLine
    : chartStyleLineCascade(linked, rawLinked, styleIndex, direct.markerStyle);
  const markerLine = direct.markerLine
    ?? (effectiveLine?.fillType === 'solid' ? effectiveLine.color
      : effectiveLine === null ? '00000000'
        : directLineAuthored ? '00000000'
        : null);
  const markerLineWidthEmu = direct.markerLineWidthEmu ?? direct.markerStyle?.lineWidthEmu
    ?? linked?.lineWidthEmu ?? null;
  const markerSize = direct.markerSize ?? chart.chartStyleMarkerSizePt;
  const markerSymbol = direct.markerSymbol ?? chart.chartStyleMarkerSymbol;
  const dataPointOverrides = direct.dataPointOverrides?.map(point => {
    if (!pointHasMarkerDetail(point)) return point;
    const directPointStyleFill = chartStyleDirectFillDecision(
      point.markerStyle, rawLinked, point.idx,
    );
    const directPointStyleLine = chartStyleDirectLineDecision(
      point.markerStyle, rawLinked, point.idx,
    );
    const pointFillAuthored = point.markerFillPaintAuthored === true
        && point.markerStyle?.fillHidden !== true
      || point.markerFill != null || point.markerFillPaint !== undefined
      || directPointStyleFill !== undefined;
    const pointLineAuthored = point.markerLine != null
      || point.markerLinePaintAuthored === true && point.markerStyle?.lineHidden !== true
      || directPointStyleLine !== undefined;
    const linkedPointFill = !pointFillAuthored
      ? directFillAuthored ? effectiveFill
          : chartStyleFillCascade(linked, rawLinked, styleIndex, direct.markerStyle)
      : undefined;
    const nextFill = point.markerFill
      ?? (directPointStyleFill?.fillType === 'solid' ? directPointStyleFill.color
        : directPointStyleFill === null ? '00000000' : undefined)
      ?? (linkedPointFill?.fillType === 'solid' ? linkedPointFill.color
        : linkedPointFill === null ? '00000000' : undefined);
    const nextFillPaint = point.markerFillPaint !== undefined
      ? point.markerFillPaint
      : directPointStyleFill?.fillType === 'gradient'
          || directPointStyleFill?.fillType === 'pattern'
          || directPointStyleFill?.fillType === 'image'
        ? directPointStyleFill
      : linkedPointFill?.fillType === 'gradient'
          || linkedPointFill?.fillType === 'pattern'
          || linkedPointFill?.fillType === 'image'
        ? linkedPointFill
        : undefined;
    const nextLine = point.markerLine
      ?? (directPointStyleLine?.fillType === 'solid' ? directPointStyleLine.color
        : directPointStyleLine === null ? '00000000' : undefined)
      ?? (pointLineAuthored ? '00000000'
        : effectiveLine?.fillType === 'solid' ? effectiveLine.color
        : effectiveLine === null ? '00000000'
        : undefined);
    const nextLineWidth = point.markerLineWidthEmu ?? point.markerStyle?.lineWidthEmu
      ?? (!pointLineAuthored && !directLineAuthored
        ? linked?.lineWidthEmu ?? undefined : undefined);
    if (nextFill === point.markerFill
      && nextFillPaint === point.markerFillPaint
      && nextLine === point.markerLine
      && nextLineWidth === point.markerLineWidthEmu) return point;
    return {
      ...point,
      markerFill: nextFill,
      markerFillPaint: nextFillPaint,
      markerFillPaintAuthored: point.markerFillPaintAuthored
        ?? (linkedPointFill !== undefined ? true : undefined),
      markerLine: nextLine,
      markerLinePaintAuthored: pointLineAuthored
        ? point.markerLinePaintAuthored
        : undefined,
      markerLineWidthEmu: nextLineWidth,
    };
  });
  if (markerFill === direct.markerFill
    && markerFillPaint === direct.markerFillPaint
    && markerFillPaintAuthored === direct.markerFillPaintAuthored
    && markerLine === direct.markerLine
    && markerLineWidthEmu === direct.markerLineWidthEmu
    && markerSize === direct.markerSize
    && markerSymbol === direct.markerSymbol
    && dataPointOverrides?.every((point, pointIndex) =>
      point === direct.dataPointOverrides?.[pointIndex]
    ) !== false) return direct;
  return {
    ...direct,
    markerFill,
    markerFillPaint,
    markerFillPaintAuthored,
    markerLine,
    markerLinePaintAuthored: directLineAuthored
      ? direct.markerLinePaintAuthored
      : undefined,
    markerLineWidthEmu,
    markerSize,
    markerSymbol,
    dataPointOverrides,
  };
}


export interface EffectiveFrameLineStyle {
  style?: ChartExStyle | null;
  color?: string | null;
  fill?: ChartModel['plotAreaLineFill'];
  widthEmu?: number | null;
  dash?: string | null;
  dashAuthored?: boolean | null;
  customDash?: ChartModel['plotAreaLineCustomDash'];
  cap?: string | null;
  join?: string | null;
  compound?: string | null;
  hidden?: boolean | null;
  paintAuthored?: boolean | null;
}


/** Merge one chart-frame outline property-by-property. Direct DrawingML paint
 * and dash choices remain authoritative; linked Chart Style geometry fills
 * only genuinely omitted properties. */
export function effectiveFrameLineStyle(
  chart: ChartModel,
  direct: EffectiveFrameLineStyle,
  linked: ChartExStyle | null | undefined,
  rawLinked: ChartExStyle | null | undefined,
  directIndex = 0,
  linkedIndex = directIndex,
): EffectiveFrameLineStyle {
  void chart;
  if (!linked) return direct;
  let { color, fill, hidden } = direct;
  const directNoLine = hidden === true
    ? chartStyleDirectNoLineDecision(rawLinked) : undefined;
  const directStyleLine = chartStyleDirectLineDecision(
    direct.style, rawLinked, directIndex,
  );
  const directPaint = fill != null || color != null
    || directNoLine !== undefined || directStyleLine !== undefined
    || direct.paintAuthored === true && hidden !== true;
  const linkedPaint = linked.lineNoStyle !== true && (linked.linePaintAuthored === true
    || linked.lineHidden === true || linked.linePaints != null || linked.lineColors != null);
  if (!directPaint) {
    const decision = chartStyleLineDecision(linked, linkedIndex);
    if (decision === null) {
      hidden = true;
    } else if (decision?.fillType === 'solid') {
      color = decision.color;
      fill = null;
      hidden = null;
    } else if (decision !== undefined) {
      fill = decision;
      color = null;
      hidden = null;
    }
  } else if (directNoLine !== undefined) {
    hidden = true;
    color = null;
    fill = null;
  } else if (fill != null || color != null) {
    hidden = null;
  } else if (fill == null && color == null) {
    const decision = directStyleLine !== undefined
      ? directStyleLine
      : direct.paintAuthored === true ? null : undefined;
    if (decision === null) hidden = true;
    else if (decision?.fillType === 'solid') {
      color = decision.color;
      fill = null;
      hidden = null;
    } else if (decision !== undefined) {
      fill = decision;
      color = null;
      hidden = null;
    }
  }
  let dash = direct.dash;
  let customDash = direct.customDash;
  let dashAuthored = direct.dashAuthored;
  if (dashAuthored !== true && dash == null && customDash == null) {
    dash = linked.lineDash;
    customDash = linked.lineCustomDash;
    dashAuthored = linked.lineDashAuthored;
  }
  return {
    color,
    fill,
    hidden,
    paintAuthored: directPaint ? true : linkedPaint ? true : direct.paintAuthored,
    widthEmu: direct.widthEmu ?? direct.style?.lineWidthEmu ?? linked.lineWidthEmu,
    dash,
    dashAuthored,
    customDash,
    cap: direct.cap ?? direct.style?.lineCap ?? linked.lineCap,
    join: direct.join ?? direct.style?.lineJoin ?? linked.lineJoin,
    compound: direct.compound ?? direct.style?.lineCompound ?? linked.lineCompound,
  };
}


export function effectiveLinkedLabelBox(
  chart: ChartModel,
  direct: ChartLabelBox | null | undefined,
  linked: ChartExStyle | null | undefined,
  rawLinked: ChartExStyle | null | undefined,
  createFromLinked: boolean,
  linkedIndex = 0,
): ChartLabelBox | undefined {
  if (!linked || (!direct && !createFromLinked)) return direct ?? undefined;
  const source = direct ?? {};
  const effectiveFill = effectiveChartLabelBoxFill(
    source, linked, rawLinked, createFromLinked, 0, linkedIndex,
  );
  const line = effectiveFrameLineStyle(chart, {
    style: source.style,
    color: source.borderColor,
    fill: source.borderFill,
    widthEmu: source.borderWidthEmu,
    dash: source.borderDash,
    dashAuthored: source.borderDashAuthored,
    customDash: source.borderCustomDash,
    cap: source.borderCap,
    join: source.borderJoin,
    compound: source.borderCompound,
    hidden: source.borderHidden,
    paintAuthored: source.borderPaintAuthored,
  }, linked, rawLinked, 0, linkedIndex);
  return {
    ...source,
    style: source.style,
    effectFallbackStyle: linked,
    effectStyleIndex: 0,
    effectFallbackIndex: linkedIndex,
    ...effectiveFill,
    borderColor: line.color ?? undefined,
    borderFill: (line.fill as ChartLabelBox['borderFill']) ?? undefined,
    borderWidthEmu: line.widthEmu ?? undefined,
    borderDash: line.dash ?? undefined,
    borderDashAuthored: line.dashAuthored ?? undefined,
    borderCustomDash: line.customDash ?? undefined,
    borderCap: line.cap ?? undefined,
    borderJoin: line.join ?? undefined,
    borderCompound: line.compound ?? undefined,
    borderHidden: line.hidden ?? undefined,
    borderPaintAuthored: line.paintAuthored ?? undefined,
  };
}


/** Merge two directly-authored label shapes property-by-property. The higher
 * precedence shape owns an authored paint/noFill choice even when that choice
 * cannot be resolved to a Canvas paint; omitted geometry continues to inherit
 * from the lower-precedence series/linked shape. */

export function chartStyleRoleDataLabels(
  chart: ChartModel,
  direct: ChartSeriesDataLabels,
  styleIndex: number,
): ChartSeriesDataLabels {
  // A transparent dLbls/spPr is ordinary label formatting, not a request for
  // Office's filled `dataLabelCallout` recipe. Select that role only when the
  // directly authored box itself has visible paint.
  const usesCalloutRole = chartLabelBoxHasVisiblePaint(direct.labelBox);
  const linked = usesCalloutRole
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  const rawLinked = usesCalloutRole
    ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
      ?? rawLinkedChartStyleRole(chart, 'dataLabel')
    : rawLinkedChartStyleRole(chart, 'dataLabel');
  // The role itself is a legitimate label-shape source. A paint-bearing
  // linked/numeric dataLabel role therefore materializes a box even when the
  // chart has no direct dLbls/spPr; an empty/no-style role still paints
  // nothing because the resulting carrier has no visible fill or outline.
  const labelBox = linked
    ? effectiveLinkedLabelBox(chart, direct.labelBox, linked, rawLinked, true, styleIndex)
    : direct.labelBox;
  const directFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, direct.fontColor, direct.fontPaintAuthored, linked, styleIndex,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic ?? undefined,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: directFontPaint ? direct.fontHidden
      : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? chart.chartTextStyle?.textRotation
      ?? linked?.textRotation ?? undefined,
    textWrap: direct.textWrap ?? chart.chartTextStyle?.textWrap ?? linked?.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? chart.chartTextStyle?.textVerticalAnchor
      ?? linked?.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? chart.chartTextStyle?.textVerticalMode
      ?? linked?.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? chart.chartTextStyle?.textLInsEmu
      ?? linked?.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? chart.chartTextStyle?.textTInsEmu
      ?? linked?.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? chart.chartTextStyle?.textRInsEmu
      ?? linked?.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? chart.chartTextStyle?.textBInsEmu
      ?? linked?.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
    labelBox,
  };
}


export function chartStyleRoleTrendlineLabel(
  chart: ChartModel,
  direct: ChartTrendline,
  styleIndex: number,
): ChartTrendline {
  const linked = chart.chartStyleRoles?.trendlineLabel;
  const rawLinked = rawLinkedChartStyleRole(chart, 'trendlineLabel');
  const directFontPaint = direct.labelFontPaintAuthored === true
    || direct.labelFontColor != null || direct.labelFontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart, direct.labelFontColor, direct.labelFontPaintAuthored, linked, styleIndex,
  );
  return {
    ...direct,
    // Unlike `dataLabelCallout`, the `trendlineLabel` role styles the generated
    // equation/R² label shape even when the chart does not carry a local spPr.
    // Materialize it before the chart-wide paint preflight so linked gradient
    // work is charged before any family starts painting.
    labelBox: linked
      ? effectiveLinkedLabelBox(chart, direct.labelBox, linked, rawLinked, true, styleIndex)
      : direct.labelBox,
    labelFontSizeHpt: direct.labelFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    labelFontBold: direct.labelFontBold ?? chart.chartTextStyle?.fontBold
      ?? linked?.fontBold ?? undefined,
    labelFontItalic: direct.labelFontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic ?? undefined,
    labelFontColor: textPaint.color ?? undefined,
    labelFontPaintAuthored: textPaint.authored,
    labelFontHidden: directFontPaint ? direct.labelFontHidden
      : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    labelFontFace: direct.labelFontFace ?? chart.chartTextStyle?.fontFace
      ?? linked?.fontFace ?? undefined,
    labelFontLanguage: direct.labelFontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage ?? undefined,
    labelFontBaseline: direct.labelFontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline ?? undefined,
    labelTextRotation: direct.labelTextRotation ?? chart.chartTextStyle?.textRotation
      ?? linked?.textRotation ?? undefined,
    labelTextWrap: direct.labelTextWrap ?? chart.chartTextStyle?.textWrap
      ?? linked?.textWrap ?? undefined,
    labelTextVerticalAnchor: direct.labelTextVerticalAnchor
      ?? chart.chartTextStyle?.textVerticalAnchor
      ?? linked?.textVerticalAnchor ?? undefined,
    labelTextVerticalMode: direct.labelTextVerticalMode ?? chart.chartTextStyle?.textVerticalMode
      ?? linked?.textVerticalMode ?? undefined,
    labelTextLInsEmu: direct.labelTextLInsEmu ?? chart.chartTextStyle?.textLInsEmu
      ?? linked?.textLInsEmu ?? undefined,
    labelTextTInsEmu: direct.labelTextTInsEmu ?? chart.chartTextStyle?.textTInsEmu
      ?? linked?.textTInsEmu ?? undefined,
    labelTextRInsEmu: direct.labelTextRInsEmu ?? chart.chartTextStyle?.textRInsEmu
      ?? linked?.textRInsEmu ?? undefined,
    labelTextBInsEmu: direct.labelTextBInsEmu ?? chart.chartTextStyle?.textBInsEmu
      ?? linked?.textBInsEmu ?? undefined,
    labelTextBodyAuthored: direct.labelTextBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
  };
}


export function chartStyleRoleDataLabelOverride(
  chart: ChartModel,
  direct: ChartDataLabelOverride,
  seriesDirect: ChartSeriesDataLabels | null | undefined,
): ChartDataLabelOverride {
  // A visible directly-authored box opts into `dataLabelCallout`; a bare or
  // transparent spPr remains an ordinary label. Applying the callout recipe
  // merely because an indexed override exists invents a white box around
  // ordinary point labels in Office styles.
  const directAndSeriesBox = mergeChartLabelBoxes(
    direct.labelBox, seriesDirect?.labelBox,
  );
  const hasCalloutShape = chartLabelBoxHasVisiblePaint(directAndSeriesBox);
  const linked = hasCalloutShape
    ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
    : chart.chartStyleRoles?.dataLabel;
  const rawLinked = hasCalloutShape
    ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
      ?? rawLinkedChartStyleRole(chart, 'dataLabel')
    : rawLinkedChartStyleRole(chart, 'dataLabel');
  const labelBox = linked
    ? effectiveLinkedLabelBox(
        chart, directAndSeriesBox, linked, rawLinked, true, direct.idx,
      )
    : directAndSeriesBox;
  const pointFontPaint = direct.fontPaintAuthored === true
    || direct.fontColor != null || direct.fontHidden === true;
  const seriesFontPaint = seriesDirect?.fontPaintAuthored === true
    || seriesDirect?.fontColor != null || seriesDirect?.fontHidden === true;
  const textPaint = effectiveInheritedChartTextPaint(
    chart,
    pointFontPaint ? direct.fontColor : seriesFontPaint ? seriesDirect?.fontColor : undefined,
    pointFontPaint ? direct.fontPaintAuthored : seriesDirect?.fontPaintAuthored,
    linked,
    direct.idx,
  );
  return {
    ...direct,
    fontSizeHpt: direct.fontSizeHpt ?? seriesDirect?.fontSizeHpt
      ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt ?? undefined,
    fontBold: direct.fontBold ?? seriesDirect?.fontBold
      ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold ?? undefined,
    fontItalic: direct.fontItalic ?? seriesDirect?.fontItalic
      ?? chart.chartTextStyle?.fontItalic ?? linked?.fontItalic ?? undefined,
    fontColor: textPaint.color ?? undefined,
    fontPaintAuthored: textPaint.authored,
    fontHidden: pointFontPaint
      ? direct.fontHidden
      : seriesFontPaint ? seriesDirect?.fontHidden
        : chart.chartTextStyle?.fontHidden ?? linked?.fontHidden ?? undefined,
    fontFace: direct.fontFace ?? seriesDirect?.fontFace
      ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace ?? undefined,
    fontLanguage: direct.fontLanguage ?? seriesDirect?.fontLanguage
      ?? chart.chartTextStyle?.fontLanguage ?? linked?.fontLanguage ?? undefined,
    fontBaseline: direct.fontBaseline ?? seriesDirect?.fontBaseline
      ?? chart.chartTextStyle?.fontBaseline ?? linked?.fontBaseline ?? undefined,
    textRotation: direct.textRotation ?? seriesDirect?.textRotation
      ?? chart.chartTextStyle?.textRotation ?? linked?.textRotation ?? undefined,
    textWrap: direct.textWrap ?? seriesDirect?.textWrap
      ?? chart.chartTextStyle?.textWrap ?? linked?.textWrap ?? undefined,
    textVerticalAnchor: direct.textVerticalAnchor ?? seriesDirect?.textVerticalAnchor
      ?? chart.chartTextStyle?.textVerticalAnchor ?? linked?.textVerticalAnchor ?? undefined,
    textVerticalMode: direct.textVerticalMode ?? seriesDirect?.textVerticalMode
      ?? chart.chartTextStyle?.textVerticalMode ?? linked?.textVerticalMode ?? undefined,
    textLInsEmu: direct.textLInsEmu ?? seriesDirect?.textLInsEmu
      ?? chart.chartTextStyle?.textLInsEmu ?? linked?.textLInsEmu ?? undefined,
    textTInsEmu: direct.textTInsEmu ?? seriesDirect?.textTInsEmu
      ?? chart.chartTextStyle?.textTInsEmu ?? linked?.textTInsEmu ?? undefined,
    textRInsEmu: direct.textRInsEmu ?? seriesDirect?.textRInsEmu
      ?? chart.chartTextStyle?.textRInsEmu ?? linked?.textRInsEmu ?? undefined,
    textBInsEmu: direct.textBInsEmu ?? seriesDirect?.textBInsEmu
      ?? chart.chartTextStyle?.textBInsEmu ?? linked?.textBInsEmu ?? undefined,
    textBodyAuthored: direct.textBodyAuthored === true
      || seriesDirect?.textBodyAuthored === true
      || chart.chartTextStyle?.textBodyAuthored === true
      || linked?.textBodyAuthored === true || undefined,
    textAlign: direct.textAlign ?? seriesDirect?.textAlign,
    labelBox,
  };
}


export function chartStyleRoleLegend(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.legend;
  const rawLinked = rawLinkedChartStyleRole(chart, 'legend');
  if (!linked && !chart.chartTextStyle) return chart;
  let legendFill = chart.legendFill;
  let legendFillColor = chart.legendFillColor;
  let legendFillHidden = chart.legendFillHidden;
  let legendFillPaintAuthored = chart.legendFillPaintAuthored;
  const directNoFill = legendFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directFillPaint = legendFill != null || legendFillColor != null
    || directNoFill !== undefined
    || chart.legendFillPaintAuthored === true && legendFillHidden !== true;
  if (!directFillPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.legendStyle);
    if (decision === null) {
      legendFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      legendFillColor = decision.color;
      legendFill = null;
      legendFillHidden = null;
    } else if (decision !== undefined) {
      legendFill = decision;
      legendFillColor = null;
      legendFillHidden = null;
    }
    if (decision !== undefined) {
      legendFillPaintAuthored = true;
    }
  }

  const legendLine = effectiveFrameLineStyle(chart, {
    style: chart.legendStyle,
    color: chart.legendLineColor,
    fill: chart.legendLineFill,
    widthEmu: chart.legendLineWidthEmu,
    dash: chart.legendLineDash,
    dashAuthored: chart.legendLineDashAuthored,
    customDash: chart.legendLineCustomDash,
    cap: chart.legendLineCap,
    join: chart.legendLineJoin,
    compound: chart.legendLineCompound,
    hidden: chart.legendLineHidden,
    paintAuthored: chart.legendLinePaintAuthored,
  }, linked, rawLinkedChartStyleRole(chart, 'legend'));
  const legendTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.legendFontColor,
    chart.legendFontPaintAuthored,
    linked,
  );
  if (legendFill === chart.legendFill
    && legendFillColor === chart.legendFillColor
    && legendFillHidden === chart.legendFillHidden
    && legendFillPaintAuthored === chart.legendFillPaintAuthored
    && legendLine.color === chart.legendLineColor
    && legendLine.fill === chart.legendLineFill
    && legendLine.widthEmu === chart.legendLineWidthEmu
    && legendLine.dash === chart.legendLineDash
    && legendLine.dashAuthored === chart.legendLineDashAuthored
    && legendLine.customDash === chart.legendLineCustomDash
    && legendLine.cap === chart.legendLineCap
    && legendLine.join === chart.legendLineJoin
    && legendLine.compound === chart.legendLineCompound
    && legendLine.hidden === chart.legendLineHidden
    && legendLine.paintAuthored === chart.legendLinePaintAuthored
    && (chart.legendFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt) === chart.legendFontSizeHpt
    && (chart.legendFontBold ?? chart.chartTextStyle?.fontBold
      ?? linked?.fontBold) === chart.legendFontBold
    && (chart.legendFontItalic ?? chart.chartTextStyle?.fontItalic
      ?? linked?.fontItalic) === chart.legendFontItalic
    && (chart.legendFontLanguage ?? chart.chartTextStyle?.fontLanguage
      ?? linked?.fontLanguage) === chart.legendFontLanguage
    && (chart.legendFontBaseline ?? chart.chartTextStyle?.fontBaseline
      ?? linked?.fontBaseline) === chart.legendFontBaseline
    && legendTextPaint.color === chart.legendFontColor
    && legendTextPaint.authored === chart.legendFontPaintAuthored
    && (chart.legendFontFace ?? chart.chartTextStyle?.fontFace
      ?? linked?.fontFace) === chart.legendFontFace) return chart;
  return {
    ...chart,
    legendFontSizeHpt: chart.legendFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
      ?? linked?.fontSizeHpt,
    legendFontBold: chart.legendFontBold ?? chart.chartTextStyle?.fontBold ?? linked?.fontBold,
    legendFontItalic: chart.legendFontItalic
      ?? chart.chartTextStyle?.fontItalic ?? linked?.fontItalic,
    legendFontLanguage: chart.legendFontLanguage
      ?? chart.chartTextStyle?.fontLanguage ?? linked?.fontLanguage,
    legendFontBaseline: chart.legendFontBaseline
      ?? chart.chartTextStyle?.fontBaseline ?? linked?.fontBaseline,
    legendFontColor: legendTextPaint.color,
    legendFontPaintAuthored: legendTextPaint.authored,
    legendFontFace: chart.legendFontFace ?? chart.chartTextStyle?.fontFace ?? linked?.fontFace,
    legendFill,
    legendFillColor,
    legendFillHidden,
    legendFillPaintAuthored,
    legendLineColor: legendLine.color,
    legendLineFill: legendLine.fill,
    legendLineWidthEmu: legendLine.widthEmu,
    legendLineDash: legendLine.dash,
    legendLineDashAuthored: legendLine.dashAuthored,
    legendLineCustomDash: legendLine.customDash,
    legendLineCap: legendLine.cap,
    legendLineJoin: legendLine.join,
    legendLineCompound: legendLine.compound,
    legendLineHidden: legendLine.hidden,
    legendLinePaintAuthored: legendLine.paintAuthored,
  };
}


/** Resolve a chart text paint as one atomic DrawingML component. Authored
 * noFill and paints that cannot be represented by Canvas become transparent;
 * neither may reveal a lower linked/numeric/default colour. */
export function effectiveChartTextPaint(
  directColor: string | null | undefined,
  directAuthored: boolean | null | undefined,
  linked: ChartExStyle | null | undefined,
  index = 0,
): { color: string | null | undefined; authored: boolean | undefined } {
  if (directAuthored === true || directColor != null) {
    return {
      color: directColor ?? '00000000',
      authored: true,
    };
  }
  if (linked && (linked.fontPaintAuthored === true
    || linked.fontColor != null || linked.fontColors != null || linked.fontHidden === true)) {
    return {
      color: linked.fontHidden === true
        ? '00000000'
        : chartStyleFontColor(linked, index) ?? '00000000',
      authored: true,
    };
  }
  return { color: directColor, authored: undefined };
}


/** Resolve the chart-wide txPr between element-local text and the style role.
 * Paint remains atomic so an authored noFill/unresolved chart default cannot
 * leak through to linked or numeric colors. */
export function effectiveInheritedChartTextPaint(
  chart: ChartModel,
  directColor: string | null | undefined,
  directAuthored: boolean | null | undefined,
  linked: ChartExStyle | null | undefined,
  index = 0,
): { color: string | null | undefined; authored: boolean | undefined } {
  const global = effectiveChartTextPaint(
    directColor, directAuthored, chart.chartTextStyle, index,
  );
  if (global.authored === true || global.color != null) return global;
  return effectiveChartTextPaint(directColor, directAuthored, linked, index);
}


export function chartStyleRolePlotArea(chart: ChartModel): ChartModel {
  // MS-ODRAWXML defines plotArea and plotArea3D as separate required style
  // entries. Do not infer one from the other in a malformed/partial sidecar;
  // direct chart formatting stays authoritative below.
  const linked = chart.threeD
    ? chart.chartStyleRoles?.plotArea3D
    : chart.chartStyleRoles?.plotArea;
  const rawLinked = rawLinkedChartStyleRole(
    chart, chart.threeD ? 'plotArea3D' : 'plotArea',
  );
  if (!linked) return chart;
  let plotAreaFill = chart.plotAreaFill;
  let plotAreaBg = chart.plotAreaBg;
  let plotAreaFillHidden = chart.plotAreaFillHidden;
  let plotAreaFillPaintAuthored = chart.plotAreaFillPaintAuthored;
  const directNoFill = plotAreaFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directPaint = ((plotAreaFill != null || plotAreaBg != null)
      && chart.plotAreaFillAutomatic !== true)
    || directNoFill !== undefined
    || chart.plotAreaFillPaintAuthored === true && plotAreaFillHidden !== true;
  if (!directPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.plotAreaStyle);
    if (decision === null) {
      plotAreaFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      plotAreaBg = decision.color;
      plotAreaFill = null;
      plotAreaFillHidden = null;
    } else if (decision !== undefined) {
      plotAreaFill = decision;
      plotAreaBg = null;
      plotAreaFillHidden = null;
    }
    if (decision !== undefined) {
      plotAreaFillPaintAuthored = true;
    }
  }

  const plotAreaLine = effectiveFrameLineStyle(chart, {
    style: chart.plotAreaStyle,
    color: chart.plotAreaLineColor,
    fill: chart.plotAreaLineFill,
    widthEmu: chart.plotAreaLineWidthEmu,
    dash: chart.plotAreaLineDash,
    dashAuthored: chart.plotAreaLineDashAuthored,
    customDash: chart.plotAreaLineCustomDash,
    cap: chart.plotAreaLineCap,
    join: chart.plotAreaLineJoin,
    compound: chart.plotAreaLineCompound,
    hidden: chart.plotAreaLineHidden,
    paintAuthored: chart.plotAreaLinePaintAuthored,
  }, linked, rawLinkedChartStyleRole(
    chart, chart.threeD ? 'plotArea3D' : 'plotArea',
  ));
  if (plotAreaFill === chart.plotAreaFill
    && plotAreaBg === chart.plotAreaBg
    && plotAreaFillHidden === chart.plotAreaFillHidden
    && plotAreaFillPaintAuthored === chart.plotAreaFillPaintAuthored
    && plotAreaLine.color === chart.plotAreaLineColor
    && plotAreaLine.fill === chart.plotAreaLineFill
    && plotAreaLine.widthEmu === chart.plotAreaLineWidthEmu
    && plotAreaLine.dash === chart.plotAreaLineDash
    && plotAreaLine.dashAuthored === chart.plotAreaLineDashAuthored
    && plotAreaLine.customDash === chart.plotAreaLineCustomDash
    && plotAreaLine.cap === chart.plotAreaLineCap
    && plotAreaLine.join === chart.plotAreaLineJoin
    && plotAreaLine.compound === chart.plotAreaLineCompound
    && plotAreaLine.hidden === chart.plotAreaLineHidden
    && plotAreaLine.paintAuthored === chart.plotAreaLinePaintAuthored) return chart;
  return {
    ...chart,
    plotAreaFill,
    plotAreaBg,
    plotAreaFillHidden,
    plotAreaFillPaintAuthored,
    plotAreaLineColor: plotAreaLine.color,
    plotAreaLineFill: plotAreaLine.fill,
    plotAreaLineWidthEmu: plotAreaLine.widthEmu,
    plotAreaLineDash: plotAreaLine.dash,
    plotAreaLineDashAuthored: plotAreaLine.dashAuthored,
    plotAreaLineCustomDash: plotAreaLine.customDash,
    plotAreaLineCap: plotAreaLine.cap,
    plotAreaLineJoin: plotAreaLine.join,
    plotAreaLineCompound: plotAreaLine.compound,
    plotAreaLineHidden: plotAreaLine.hidden,
    plotAreaLinePaintAuthored: plotAreaLine.paintAuthored,
  };
}


export function chartStyleRoleChartArea(chart: ChartModel): ChartModel {
  const linked = chart.chartStyleRoles?.chartArea;
  const rawLinked = rawLinkedChartStyleRole(chart, 'chartArea');
  if (!linked) return chart;
  let chartFill = chart.chartFill;
  let chartBg = chart.chartBg;
  let chartFillHidden = chart.chartFillHidden;
  let chartFillPaintAuthored = chart.chartFillPaintAuthored;
  const directNoFill = chartFillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directPaint = chartFill != null || directNoFill !== undefined
    || chart.chartFillPaintAuthored === true && chartFillHidden !== true;
  if (!directPaint) {
    const decision = chartStyleFillCascade(linked, rawLinked, 0, chart.chartAreaStyle);
    if (decision === null) {
      chartFill = null;
      chartBg = null;
      chartFillHidden = true;
    } else if (decision?.fillType === 'solid') {
      chartBg = decision.color;
      chartFill = null;
      chartFillHidden = null;
    } else if (decision !== undefined) {
      chartFill = decision;
      chartBg = null;
      chartFillHidden = null;
    }
    if (decision !== undefined) {
      chartFillPaintAuthored = true;
    }
  }

  const chartBorder = effectiveFrameLineStyle(chart, {
    style: chart.chartAreaStyle,
    color: chart.chartBorderColor,
    fill: chart.chartBorderLineFill,
    widthEmu: chart.chartBorderWidthEmu,
    dash: chart.chartBorderDash,
    dashAuthored: chart.chartBorderDashAuthored,
    customDash: chart.chartBorderCustomDash,
    cap: chart.chartBorderCap,
    join: chart.chartBorderJoin,
    compound: chart.chartBorderCompound,
    hidden: chart.chartBorderHidden,
    paintAuthored: chart.chartBorderPaintAuthored,
  }, linked, rawLinkedChartStyleRole(chart, 'chartArea'));
  if (chartFill === chart.chartFill
    && chartBg === chart.chartBg
    && chartFillHidden === chart.chartFillHidden
    && chartFillPaintAuthored === chart.chartFillPaintAuthored
    && chartBorder.color === chart.chartBorderColor
    && chartBorder.fill === chart.chartBorderLineFill
    && chartBorder.widthEmu === chart.chartBorderWidthEmu
    && chartBorder.dash === chart.chartBorderDash
    && chartBorder.dashAuthored === chart.chartBorderDashAuthored
    && chartBorder.customDash === chart.chartBorderCustomDash
    && chartBorder.cap === chart.chartBorderCap
    && chartBorder.join === chart.chartBorderJoin
    && chartBorder.compound === chart.chartBorderCompound
    && chartBorder.hidden === chart.chartBorderHidden
    && chartBorder.paintAuthored === chart.chartBorderPaintAuthored) return chart;
  return {
    ...chart,
    chartFill,
    chartBg,
    chartFillHidden,
    chartFillPaintAuthored,
    chartBorderColor: chartBorder.color,
    chartBorderLineFill: chartBorder.fill,
    chartBorderWidthEmu: chartBorder.widthEmu,
    chartBorderDash: chartBorder.dash,
    chartBorderDashAuthored: chartBorder.dashAuthored,
    chartBorderCustomDash: chartBorder.customDash,
    chartBorderCap: chartBorder.cap,
    chartBorderJoin: chartBorder.join,
    chartBorderCompound: chartBorder.compound,
    chartBorderHidden: chartBorder.hidden,
    chartBorderPaintAuthored: chartBorder.paintAuthored,
  };
}


/** Materialize the linked decoration roles that an optional family renderer
 * consumes directly from `ChartSeries`. Keeping this projection in core means
 * the 2-D, 3-D, DOCX, XLSX, and PPTX paths receive one effective precedence
 * result without teaching an optional renderer about package sidecars. */
export function withOfficeStyleRasterLineFloor(chart: ChartModel, ptToPx: number): ChartModel {
  const roles = chart.chartStyleRoles;
  if (!roles || !(Number.isFinite(ptToPx) && ptToPx > 0)) return chart;
  // Office keeps the authored/theme width in its vector output (for example,
  // classic Style 2 emits a 6,350 EMU / 0.5pt black rule), then stroke-adjusts
  // that vector to one opaque device pixel when an axis or gridline would
  // otherwise land below a pixel. Canvas instead alpha-antialiases the
  // subpixel stroke into a grey rule. Apply the observed raster floor only to
  // those evidenced style roles in this render projection: direct `<a:ln w>`
  // remains exact, other role families stay untouched, the source model keeps
  // its ECMA-376 width, and zoomed widths naturally exceed the floor.
  const minimumWidthEmu = EMU_PER_PT / ptToPx;
  let changed = false;
  const strokeAdjustedRoles = new Set<ChartStyleRole>([
    'categoryAxis', 'seriesAxis', 'valueAxis', 'gridlineMajor', 'gridlineMinor',
  ]);
  const adjusted = Object.fromEntries(Object.entries(roles).map(([role, style]) => {
    if (!strokeAdjustedRoles.has(role as ChartStyleRole)) return [role, style];
    if (!style || style.lineWidthEmu == null
      || !Number.isFinite(style.lineWidthEmu)
      || style.lineWidthEmu <= 0
      || style.lineWidthEmu >= minimumWidthEmu) return [role, style];
    changed = true;
    return [role, { ...style, lineWidthEmu: minimumWidthEmu }];
  })) as typeof roles;
  return changed ? { ...chart, chartStyleRoles: adjusted } : chart;
}


export function applyLinkedChartStyleRoles(chart: ChartModel, ptToPx: number): ChartModel {
  chart = withOfficeStyleRasterLineFloor(chart, ptToPx);
  if (!chart.chartStyleRoles?.errorBar
    && !chart.chartStyleRoles?.leaderLine
    && !chart.chartStyleRoles?.trendline
    && !chart.chartStyleRoles?.trendlineLabel
    && !chart.chartStyleRoles?.dataLabel
    && !chart.chartStyleRoles?.dataLabelCallout
    && !chart.chartStyleRoles?.dataTable
    && !chart.chartStyleRoles?.gridlineMajor
    && !chart.chartStyleRoles?.gridlineMinor
    && !chart.chartStyleRoles?.categoryAxis
    && !chart.chartStyleRoles?.valueAxis
    && !chart.chartStyleRoles?.seriesAxis
    && !chart.chartStyleRoles?.dataPointMarker
    && !chart.chartStyleRoles?.legend
    && !chart.chartStyleRoles?.plotArea
    && !chart.chartStyleRoles?.plotArea3D
    && !chart.chartStyleRoles?.chartArea
    && !chart.chartStyleRoles?.title
    && !chart.chartStyleRoles?.axisTitle
    && chart.chartTextStyle == null
    && chart.chartStyleMarkerSizePt == null
    && chart.chartStyleMarkerSymbol == null) {
    return chart;
  }
  let changed = false;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const series = chart.series.map((sourceItem, seriesIndex) => {
    const item = chartStyleRoleMarker(
      chart, sourceItem, seriesIndex, chart.series.length, plotGroupBySeries[seriesIndex],
    );
    changed ||= item !== sourceItem;
    const errBars = chart.chartStyleRoles?.errorBar ? item.errBars?.map(errorBar => {
      const effective = chartStyleRoleErrorBar(chart, errorBar);
      changed ||= effective.color !== errorBar.color
        || effective.lineWidthEmu !== errorBar.lineWidthEmu
        || effective.dash !== errorBar.dash
        || effective.hidden !== errorBar.hidden;
      return effective;
    }) : item.errBars;
    let seriesDataLabels = item.seriesDataLabels;
    if (seriesDataLabels
      && (chart.chartTextStyle
        || chart.chartStyleRoles?.dataLabel || chart.chartStyleRoles?.dataLabelCallout)) {
      const effective = chartStyleRoleDataLabels(
        chart,
        seriesDataLabels,
        chartExSeriesFormatIndex(item, seriesIndex),
      );
      changed ||= effective !== seriesDataLabels;
      seriesDataLabels = effective;
    }
    const dataLabelOverrides = (chart.chartTextStyle || chart.chartStyleRoles?.dataLabelCallout
      || chart.chartStyleRoles?.dataLabel)
      ? item.dataLabelOverrides?.map(override => {
          const effective = chartStyleRoleDataLabelOverride(
            chart,
            override,
            sourceItem.seriesDataLabels,
          );
          changed ||= effective !== override;
          return effective;
        })
      : item.dataLabelOverrides;
    if (seriesDataLabels && chart.chartStyleRoles?.leaderLine) {
      const effective = chartStyleRoleLeaderLine(chart, seriesDataLabels);
      const merged = {
        ...seriesDataLabels,
        leaderLineColor: effective.color ?? undefined,
        leaderLineWidthEmu: effective.widthEmu ?? undefined,
        leaderLineDash: effective.dash ?? undefined,
        leaderLineHidden: effective.hidden ?? undefined,
        leaderLinePaintAuthored: effective.paintAuthored,
      };
      changed ||= merged.leaderLineColor !== seriesDataLabels.leaderLineColor
        || merged.leaderLineWidthEmu !== seriesDataLabels.leaderLineWidthEmu
        || merged.leaderLineDash !== seriesDataLabels.leaderLineDash
        || merged.leaderLineHidden !== seriesDataLabels.leaderLineHidden
        || merged.leaderLinePaintAuthored !== seriesDataLabels.leaderLinePaintAuthored;
      seriesDataLabels = merged;
    }
    const trendLines = (chart.chartStyleRoles?.trendline || chart.chartTextStyle
      || chart.chartStyleRoles?.trendlineLabel)
      ? item.trendLines?.map(trendline => {
      let effective = chart.chartStyleRoles?.trendline
        ? chartStyleRoleTrendline(chart, trendline)
        : trendline;
      if (chart.chartTextStyle || chart.chartStyleRoles?.trendlineLabel) {
        effective = chartStyleRoleTrendlineLabel(
          chart,
          effective,
          chartExSeriesFormatIndex(item, seriesIndex),
        );
      }
      changed ||= effective.lineColor !== trendline.lineColor
        || effective.lineWidthEmu !== trendline.lineWidthEmu
        || effective.lineDash !== trendline.lineDash
        || effective.lineHidden !== trendline.lineHidden
        || effective !== trendline;
      return effective;
    }) : item.trendLines;
    if (errBars === item.errBars
      && seriesDataLabels === item.seriesDataLabels
      && dataLabelOverrides === item.dataLabelOverrides
      && trendLines === item.trendLines) return item;
    return { ...item, errBars, seriesDataLabels, dataLabelOverrides, trendLines };
  });
  let dataTable = chart.dataTable;
  if (dataTable && (chart.chartStyleRoles?.dataTable || chart.chartTextStyle)) {
    const effective = chartStyleRoleDataTable(chart, dataTable);
    changed ||= effective !== dataTable;
    dataTable = effective;
  }
  const valMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.valAxisMajorGridlines,
    chart.valAxisGridlineColor, chart.valAxisGridlineWidthEmu, chart.valAxisGridlineDash,
    chart.valAxisGridlinePaintAuthored,
    chart.valAxisMajorGridlineStyle,
  );
  const catMajor = chartStyleRoleGridline(
    chart, 'gridlineMajor', chart.catAxisMajorGridlines,
    chart.catAxisGridlineColor, chart.catAxisGridlineWidthEmu, chart.catAxisGridlineDash,
    chart.catAxisGridlinePaintAuthored,
    chart.catAxisMajorGridlineStyle,
  );
  const valMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.valAxisMinorGridlines,
    chart.valAxisMinorGridlineColor,
    chart.valAxisMinorGridlineWidthEmu,
    chart.valAxisMinorGridlineDash,
    chart.valAxisMinorGridlinePaintAuthored,
    chart.valAxisMinorGridlineStyle,
  );
  const catMinor = chartStyleRoleGridline(
    chart, 'gridlineMinor', chart.catAxisMinorGridlines,
    chart.catAxisMinorGridlineColor,
    chart.catAxisMinorGridlineWidthEmu,
    chart.catAxisMinorGridlineDash,
    chart.catAxisMinorGridlinePaintAuthored,
    chart.catAxisMinorGridlineStyle,
  );
  const secondaryValGridlines = chartStyleRoleSecondaryGridlines(chart, chart.secondaryValAxis);
  const secondaryCatGridlines = chartStyleRoleSecondaryGridlines(chart, chart.secondaryCatAxis);
  const secondaryValAxis = chartStyleRoleSecondaryAxisLine(
    chart, secondaryValGridlines, 'valueAxis',
  );
  const secondaryCatAxis = chartStyleRoleSecondaryAxisLine(
    chart, secondaryCatGridlines, 'categoryAxis',
  );
  const catAxisLine = chartStyleRoleAxisLine(
    chart, 'categoryAxis',
    chart.catAxisLineColor, chart.catAxisLineWidthEmu, chart.catAxisLineDash,
    chart.catAxisLineHidden,
    chart.catAxisLinePaintAuthored,
    chart.catAxisStyle,
  );
  const valAxisLine = chartStyleRoleAxisLine(
    chart, 'valueAxis',
    chart.valAxisLineColor, chart.valAxisLineWidthEmu, chart.valAxisLineDash,
    chart.valAxisLineHidden,
    chart.valAxisLinePaintAuthored,
    chart.valAxisStyle,
  );
  const catAxisLineHidden = chartAxisLineIsHidden(catAxisLine);
  const valAxisLineHidden = chartAxisLineIsHidden(valAxisLine);
  const catAxisStyle = chart.chartStyleRoles?.categoryAxis;
  const valAxisStyle = chart.chartStyleRoles?.valueAxis;
  const titleStyle = chart.chartStyleRoles?.title;
  const axisTitleStyle = chart.chartStyleRoles?.axisTitle;
  const dataLabelStyle = chart.chartStyleRoles?.dataLabel;
  const catAxisFontSizeHpt = chart.catAxisFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? catAxisStyle?.fontSizeHpt ?? null;
  const catAxisFontBold = chart.catAxisFontBold ?? chart.chartTextStyle?.fontBold
    ?? catAxisStyle?.fontBold;
  const catAxisFontItalic = chart.catAxisFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? catAxisStyle?.fontItalic;
  const catAxisTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.catAxisFontColor, chart.catAxisFontPaintAuthored, catAxisStyle,
  );
  const catAxisFontColor = catAxisTextPaint.color;
  const catAxisFontFace = chart.catAxisFontFace ?? chart.chartTextStyle?.fontFace
    ?? catAxisStyle?.fontFace;
  const valAxisFontSizeHpt = chart.valAxisFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? valAxisStyle?.fontSizeHpt ?? null;
  const valAxisFontBold = chart.valAxisFontBold ?? chart.chartTextStyle?.fontBold
    ?? valAxisStyle?.fontBold;
  const valAxisFontItalic = chart.valAxisFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? valAxisStyle?.fontItalic;
  const valAxisTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.valAxisFontColor, chart.valAxisFontPaintAuthored, valAxisStyle,
  );
  const valAxisFontColor = valAxisTextPaint.color;
  const valAxisFontFace = chart.valAxisFontFace ?? chart.chartTextStyle?.fontFace
    ?? valAxisStyle?.fontFace;
  const titleFontSizeHpt = chart.titleFontSizeHpt ?? chart.chartTextStyle?.fontSizeHpt
    ?? titleStyle?.fontSizeHpt ?? null;
  const titleFontBold = chart.titleFontBold ?? chart.chartTextStyle?.fontBold
    ?? titleStyle?.fontBold;
  const titleFontItalic = chart.titleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? titleStyle?.fontItalic;
  const titleFontLanguage = chart.titleFontLanguage ?? chart.chartTextStyle?.fontLanguage
    ?? titleStyle?.fontLanguage;
  const titleFontBaseline = chart.titleFontBaseline ?? chart.chartTextStyle?.fontBaseline
    ?? titleStyle?.fontBaseline;
  const titleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.titleFontColor, chart.titleFontPaintAuthored, titleStyle,
  );
  const titleFontColor = titleTextPaint.color ?? null;
  const titleFontFace = chart.titleFontFace ?? chart.chartTextStyle?.fontFace
    ?? titleStyle?.fontFace ?? null;
  const catAxisTitleFontSizeHpt = chart.catAxisTitleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? axisTitleStyle?.fontSizeHpt;
  const catAxisTitleFontBold = chart.catAxisTitleFontBold ?? chart.chartTextStyle?.fontBold
    ?? axisTitleStyle?.fontBold;
  const catAxisTitleFontItalic = chart.catAxisTitleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? axisTitleStyle?.fontItalic;
  const catAxisTitleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.catAxisTitleFontColor, chart.catAxisTitleFontPaintAuthored, axisTitleStyle,
  );
  const catAxisTitleFontColor = catAxisTitleTextPaint.color;
  const catAxisTitleFontFace = chart.catAxisTitleFontFace ?? chart.chartTextStyle?.fontFace
    ?? axisTitleStyle?.fontFace;
  const valAxisTitleFontSizeHpt = chart.valAxisTitleFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? axisTitleStyle?.fontSizeHpt;
  const valAxisTitleFontBold = chart.valAxisTitleFontBold ?? chart.chartTextStyle?.fontBold
    ?? axisTitleStyle?.fontBold;
  const valAxisTitleFontItalic = chart.valAxisTitleFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? axisTitleStyle?.fontItalic;
  const valAxisTitleTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.valAxisTitleFontColor, chart.valAxisTitleFontPaintAuthored, axisTitleStyle,
  );
  const valAxisTitleFontColor = valAxisTitleTextPaint.color;
  const valAxisTitleFontFace = chart.valAxisTitleFontFace ?? chart.chartTextStyle?.fontFace
    ?? axisTitleStyle?.fontFace;
  const dataLabelFontSizeHpt = chart.dataLabelFontSizeHpt
    ?? chart.chartTextStyle?.fontSizeHpt ?? dataLabelStyle?.fontSizeHpt ?? null;
  const dataLabelFontBold = chart.dataLabelFontBold ?? chart.chartTextStyle?.fontBold
    ?? dataLabelStyle?.fontBold;
  const dataLabelFontItalic = chart.dataLabelFontItalic ?? chart.chartTextStyle?.fontItalic
    ?? dataLabelStyle?.fontItalic;
  const dataLabelFontLanguage = chart.dataLabelFontLanguage
    ?? chart.chartTextStyle?.fontLanguage ?? dataLabelStyle?.fontLanguage;
  const dataLabelFontBaseline = chart.dataLabelFontBaseline
    ?? chart.chartTextStyle?.fontBaseline ?? dataLabelStyle?.fontBaseline;
  const dataLabelTextPaint = effectiveInheritedChartTextPaint(
    chart,
    chart.dataLabelFontColor, chart.dataLabelFontPaintAuthored, dataLabelStyle,
  );
  const dataLabelFontColor = dataLabelTextPaint.color;
  const dataLabelFontFace = chart.dataLabelFontFace ?? chart.chartTextStyle?.fontFace
    ?? dataLabelStyle?.fontFace;
  changed ||= valMajor.visible !== chart.valAxisMajorGridlines
    || valMajor.color !== chart.valAxisGridlineColor
    || valMajor.widthEmu !== chart.valAxisGridlineWidthEmu
    || valMajor.dash !== chart.valAxisGridlineDash
    || valMajor.paintAuthored !== chart.valAxisGridlinePaintAuthored
    || catMajor.visible !== chart.catAxisMajorGridlines
    || catMajor.color !== chart.catAxisGridlineColor
    || catMajor.widthEmu !== chart.catAxisGridlineWidthEmu
    || catMajor.dash !== chart.catAxisGridlineDash
    || catMajor.paintAuthored !== chart.catAxisGridlinePaintAuthored
    || valMinor.visible !== chart.valAxisMinorGridlines
    || valMinor.color !== chart.valAxisMinorGridlineColor
    || valMinor.widthEmu !== chart.valAxisMinorGridlineWidthEmu
    || valMinor.dash !== chart.valAxisMinorGridlineDash
    || valMinor.paintAuthored !== chart.valAxisMinorGridlinePaintAuthored
    || catMinor.visible !== chart.catAxisMinorGridlines
    || catMinor.color !== chart.catAxisMinorGridlineColor
    || catMinor.widthEmu !== chart.catAxisMinorGridlineWidthEmu
    || catMinor.dash !== chart.catAxisMinorGridlineDash
    || catMinor.paintAuthored !== chart.catAxisMinorGridlinePaintAuthored
    || secondaryValAxis !== chart.secondaryValAxis
    || secondaryCatAxis !== chart.secondaryCatAxis
    || catAxisLine.color !== chart.catAxisLineColor
    || catAxisLine.widthEmu !== chart.catAxisLineWidthEmu
    || catAxisLine.dash !== chart.catAxisLineDash
    || catAxisLineHidden !== chart.catAxisLineHidden
    || catAxisLine.paintAuthored !== chart.catAxisLinePaintAuthored
    || valAxisLine.color !== chart.valAxisLineColor
    || valAxisLine.widthEmu !== chart.valAxisLineWidthEmu
    || valAxisLine.dash !== chart.valAxisLineDash
    || valAxisLineHidden !== chart.valAxisLineHidden
    || valAxisLine.paintAuthored !== chart.valAxisLinePaintAuthored
    || catAxisFontSizeHpt !== chart.catAxisFontSizeHpt
    || catAxisFontBold !== chart.catAxisFontBold
    || catAxisFontItalic !== chart.catAxisFontItalic
    || catAxisFontColor !== chart.catAxisFontColor
    || catAxisTextPaint.authored !== chart.catAxisFontPaintAuthored
    || catAxisFontFace !== chart.catAxisFontFace
    || valAxisFontSizeHpt !== chart.valAxisFontSizeHpt
    || valAxisFontBold !== chart.valAxisFontBold
    || valAxisFontItalic !== chart.valAxisFontItalic
    || valAxisFontColor !== chart.valAxisFontColor
    || valAxisTextPaint.authored !== chart.valAxisFontPaintAuthored
    || valAxisFontFace !== chart.valAxisFontFace
    || titleFontSizeHpt !== chart.titleFontSizeHpt
    || titleFontBold !== chart.titleFontBold
    || titleFontItalic !== chart.titleFontItalic
    || titleFontLanguage !== chart.titleFontLanguage
    || titleFontBaseline !== chart.titleFontBaseline
    || titleFontColor !== chart.titleFontColor
    || titleTextPaint.authored !== chart.titleFontPaintAuthored
    || titleFontFace !== chart.titleFontFace
    || catAxisTitleFontSizeHpt !== chart.catAxisTitleFontSizeHpt
    || catAxisTitleFontBold !== chart.catAxisTitleFontBold
    || catAxisTitleFontItalic !== chart.catAxisTitleFontItalic
    || catAxisTitleFontColor !== chart.catAxisTitleFontColor
    || catAxisTitleTextPaint.authored !== chart.catAxisTitleFontPaintAuthored
    || catAxisTitleFontFace !== chart.catAxisTitleFontFace
    || valAxisTitleFontSizeHpt !== chart.valAxisTitleFontSizeHpt
    || valAxisTitleFontBold !== chart.valAxisTitleFontBold
    || valAxisTitleFontItalic !== chart.valAxisTitleFontItalic
    || valAxisTitleFontColor !== chart.valAxisTitleFontColor
    || valAxisTitleTextPaint.authored !== chart.valAxisTitleFontPaintAuthored
    || valAxisTitleFontFace !== chart.valAxisTitleFontFace
    || dataLabelFontSizeHpt !== chart.dataLabelFontSizeHpt
    || dataLabelFontBold !== chart.dataLabelFontBold
    || dataLabelFontItalic !== chart.dataLabelFontItalic
    || dataLabelFontLanguage !== chart.dataLabelFontLanguage
    || dataLabelFontBaseline !== chart.dataLabelFontBaseline
    || dataLabelFontColor !== chart.dataLabelFontColor
    || dataLabelTextPaint.authored !== chart.dataLabelFontPaintAuthored
    || dataLabelFontFace !== chart.dataLabelFontFace;
  const effective = changed ? {
    ...chart,
    series,
    dataTable,
    valAxisMajorGridlines: valMajor.visible,
    valAxisGridlineColor: valMajor.color,
    valAxisGridlineWidthEmu: valMajor.widthEmu,
    valAxisGridlineDash: valMajor.dash,
    valAxisGridlinePaintAuthored: valMajor.paintAuthored,
    catAxisMajorGridlines: catMajor.visible,
    catAxisGridlineColor: catMajor.color,
    catAxisGridlineWidthEmu: catMajor.widthEmu,
    catAxisGridlineDash: catMajor.dash,
    catAxisGridlinePaintAuthored: catMajor.paintAuthored,
    valAxisMinorGridlines: valMinor.visible,
    valAxisMinorGridlineColor: valMinor.color,
    valAxisMinorGridlineWidthEmu: valMinor.widthEmu,
    valAxisMinorGridlineDash: valMinor.dash,
    valAxisMinorGridlinePaintAuthored: valMinor.paintAuthored,
    catAxisMinorGridlines: catMinor.visible,
    catAxisMinorGridlineColor: catMinor.color,
    catAxisMinorGridlineWidthEmu: catMinor.widthEmu,
    catAxisMinorGridlineDash: catMinor.dash,
    catAxisMinorGridlinePaintAuthored: catMinor.paintAuthored,
    secondaryValAxis,
    secondaryCatAxis,
    catAxisLineColor: catAxisLine.color,
    catAxisLineWidthEmu: catAxisLine.widthEmu,
    catAxisLineDash: catAxisLine.dash,
    catAxisLineHidden,
    catAxisLinePaintAuthored: catAxisLine.paintAuthored,
    valAxisLineColor: valAxisLine.color,
    valAxisLineWidthEmu: valAxisLine.widthEmu,
    valAxisLineDash: valAxisLine.dash,
    valAxisLineHidden,
    valAxisLinePaintAuthored: valAxisLine.paintAuthored,
    catAxisFontSizeHpt,
    catAxisFontBold,
    catAxisFontItalic,
    catAxisFontColor,
    catAxisFontPaintAuthored: catAxisTextPaint.authored,
    catAxisFontFace,
    valAxisFontSizeHpt,
    valAxisFontBold,
    valAxisFontItalic,
    valAxisFontColor,
    valAxisFontPaintAuthored: valAxisTextPaint.authored,
    valAxisFontFace,
    titleFontSizeHpt,
    titleFontBold,
    titleFontItalic,
    titleFontLanguage,
    titleFontBaseline,
    titleFontColor,
    titleFontPaintAuthored: titleTextPaint.authored,
    titleFontFace,
    catAxisTitleFontSizeHpt,
    catAxisTitleFontBold,
    catAxisTitleFontItalic,
    catAxisTitleFontColor,
    catAxisTitleFontPaintAuthored: catAxisTitleTextPaint.authored,
    catAxisTitleFontFace,
    valAxisTitleFontSizeHpt,
    valAxisTitleFontBold,
    valAxisTitleFontItalic,
    valAxisTitleFontColor,
    valAxisTitleFontPaintAuthored: valAxisTitleTextPaint.authored,
    valAxisTitleFontFace,
    dataLabelFontSizeHpt,
    dataLabelFontBold,
    dataLabelFontItalic,
    dataLabelFontLanguage,
    dataLabelFontBaseline,
    dataLabelFontColor,
    dataLabelFontPaintAuthored: dataLabelTextPaint.authored,
    dataLabelFontFace,
  } : chart;
  return chartStyleRoleLegend(chartStyleRolePlotArea(chartStyleRoleChartArea(
    chartStyleRoleSeriesAxis(effective),
  )));
}
