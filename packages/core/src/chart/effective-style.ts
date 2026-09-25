import type {
  ChartExElementStyle,
  ChartModel,
  ChartStyleRole,
} from '../types/chart.js';

const TEXT_GEOMETRY_KEYS = [
  'fontSizeHpt', 'fontBold', 'fontItalic', 'fontFace', 'fontLanguage',
  'fontBaseline', 'textRotation',
  'textWrap', 'textVerticalAnchor', 'textVerticalMode', 'textLInsEmu',
  'textTInsEmu', 'textRInsEmu', 'textBInsEmu', 'textBodyAuthored',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const TEXT_PAINT_KEYS = [
  'fontColor', 'fontColors', 'fontColorIndex', 'fontFormattingIndices',
  'fontPaintAuthored', 'fontHidden',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const FILL_KEYS = [
  'fillPaints', 'fillColors', 'fillHidden', 'fillPaintAuthored',
  'fillNoStyle', 'fillColorIndex', 'fillFormattingIndices',
  'fillSemanticFallbackIndices',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const LINE_PAINT_KEYS = [
  'lineColors', 'linePaints', 'linePaintAuthored', 'lineHidden',
  'lineNoStyle',
  'lineColorIndex', 'lineFormattingIndices',
  'lineSemanticFallbackIndices',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const MODIFIER_KEYS = [
  'shapePropertiesPresent', 'allowNoFillOverride', 'allowNoLineOverride',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const LINE_GEOMETRY_KEYS = [
  'lineWidthEmu', 'lineCap', 'lineJoin', 'lineCompound',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

const LINE_DASH_KEYS = [
  'lineDash', 'lineDashAuthored', 'lineCustomDash',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

export type ChartDashStyle = Pick<
  ChartExElementStyle,
  'lineDash' | 'lineDashAuthored' | 'lineCustomDash'
>;

interface ChartStyleIndexCache {
  series: WeakMap<object, number>;
  groupBySeries: Int32Array;
}

let activeChartStyleIndexCaches: WeakMap<ChartModel, ChartStyleIndexCache> | undefined;

/** Scope source-series/group indexes to one synchronous render. The public
 * model may be mutated and reused by an application between render calls. */
export function withChartStyleIndexCache<T>(run: () => T): T {
  if (activeChartStyleIndexCaches) return run();
  activeChartStyleIndexCaches = new WeakMap();
  try {
    return run();
  } finally {
    activeChartStyleIndexCaches = undefined;
  }
}

function chartStyleIndexCache(chart: ChartModel): ChartStyleIndexCache {
  const cached = activeChartStyleIndexCaches?.get(chart);
  if (cached) return cached;
  const series = new WeakMap<object, number>();
  for (let index = 0; index < chart.series.length; index++) {
    series.set(chart.series[index]!, index);
  }
  const groupBySeries = new Int32Array(chart.series.length);
  groupBySeries.fill(-1);
  for (let groupIndex = 0; groupIndex < (chart.plotGroups?.length ?? 0); groupIndex++) {
    const group = chart.plotGroups![groupIndex]!;
    const end = Math.min(chart.series.length, group.seriesStart + group.seriesCount);
    for (let index = Math.max(0, group.seriesStart); index < end; index++) {
      groupBySeries[index] = groupIndex;
    }
  }
  const result = { series, groupBySeries };
  activeChartStyleIndexCaches?.set(chart, result);
  return result;
}

/** O(1) source-series lookup shared by per-datum style hot paths. */
export function chartSeriesSourceIndex(
  chart: ChartModel,
  series: ChartModel['series'][number],
  fallback = -1,
): number {
  return chartStyleIndexCache(chart).series.get(series) ?? fallback;
}

function chartStyleGroupIndex(chart: ChartModel, seriesIndex: number): number {
  return seriesIndex >= 0 && seriesIndex < chart.series.length
    ? chartStyleIndexCache(chart).groupBySeries[seriesIndex] ?? -1
    : -1;
}

/** O(1) owning plot-group lookup after one bounded chart indexing pass. */
export function chartPlotGroupForSeries(
  chart: ChartModel,
  seriesIndex: number,
): NonNullable<ChartModel['plotGroups']>[number] | undefined {
  const groupIndex = chartStyleGroupIndex(chart, seriesIndex);
  return groupIndex >= 0 ? chart.plotGroups?.[groupIndex] : undefined;
}

/** Select DrawingML dash as one atomic choice. A preset and custom dash cannot
 * inherit independently across direct, linked, and numeric layers. */
export function chartStyleDashChoice(
  ...layers: Array<ChartDashStyle | null | undefined>
): ChartDashStyle | undefined {
  for (const layer of layers) {
    if (layer != null && (layer.lineDashAuthored === true
      || layer.lineDash != null || layer.lineCustomDash != null)) return layer;
  }
  return undefined;
}

const EFFECT_KEYS = [
  'shadows', 'innerShadows', 'glows', 'softEdges', 'reflections',
  'effectAuthored', 'effectNoStyle', 'effectUnsupported',
  'effectFormattingIndices', 'effectColorIndex',
] as const satisfies ReadonlyArray<keyof ChartExElementStyle>;

function copyPresent(
  target: ChartExElementStyle,
  source: ChartExElementStyle | null | undefined,
  keys: ReadonlyArray<keyof ChartExElementStyle>,
): void {
  if (!source) return;
  const targetRecord = target as Record<string, unknown>;
  const sourceRecord = source as Record<string, unknown>;
  for (const key of keys) {
    const value = sourceRecord[key];
    if (value !== undefined && value !== null) targetRecord[key] = value;
  }
}

function hasPresent(
  style: ChartExElementStyle,
  keys: ReadonlyArray<keyof ChartExElementStyle>,
): boolean {
  const record = style as Record<string, unknown>;
  return keys.some(key => record[key] !== undefined && record[key] !== null);
}

/**
 * Compose one linked Office 2013+ Chart Style role over its ECMA-376 numeric
 * style fallback. Fill, line, effect, and text are independent style components:
 * omitting a component preserves the numeric default, an explicit no-fill
 * suppresses it, and the linked `NoStyle` sentinel falls through to it.
 */
export function effectiveChartStyleRole(
  numeric: ChartExElementStyle | null | undefined,
  linked: ChartExElementStyle | null | undefined,
): ChartExElementStyle | undefined {
  if (!numeric) return linked ?? undefined;
  if (!linked) return numeric;

  const effective: ChartExElementStyle = {};
  // Text typography inherits property-by-property, but text paint is an
  // atomic DrawingML choice. An authored linked noFill or unresolved paint
  // must not expose a less-specific numeric font color.
  copyPresent(effective, numeric, TEXT_GEOMETRY_KEYS);
  copyPresent(effective, linked, TEXT_GEOMETRY_KEYS);
  copyPresent(effective, numeric, TEXT_PAINT_KEYS);
  if (hasPresent(linked, TEXT_PAINT_KEYS)) {
    for (const key of TEXT_PAINT_KEYS) delete (effective as Record<string, unknown>)[key];
    copyPresent(effective, linked, TEXT_PAINT_KEYS);
  }

  copyPresent(effective, numeric, FILL_KEYS);
  if (linked.fillNoStyle !== true && hasPresent(linked, FILL_KEYS)) {
    for (const key of FILL_KEYS) delete (effective as Record<string, unknown>)[key];
    copyPresent(effective, linked, FILL_KEYS);
  }

  copyPresent(effective, numeric, LINE_PAINT_KEYS);
  copyPresent(effective, numeric, LINE_GEOMETRY_KEYS);
  copyPresent(effective, numeric, LINE_DASH_KEYS);
  if (linked.lineNoStyle !== true) {
    // Line paint is one DrawingML choice. A linked paint atomically replaces
    // the numeric one; NoStyle leaves numeric paint intact.
    if (hasPresent(linked, LINE_PAINT_KEYS)) {
      for (const key of LINE_PAINT_KEYS) delete (effective as Record<string, unknown>)[key];
      copyPresent(effective, linked, LINE_PAINT_KEYS);
    }
  }
  // Width/dash/cap/join are independent from paint. In particular, an
  // lnRef idx=0 role may still carry local a:ln geometry which must overlay
  // numeric geometry while its paint falls through.
  copyPresent(effective, linked, LINE_GEOMETRY_KEYS);
  if (hasPresent(linked, LINE_DASH_KEYS)) {
    for (const key of LINE_DASH_KEYS) delete (effective as Record<string, unknown>)[key];
    copyPresent(effective, linked, LINE_DASH_KEYS);
  }

  copyPresent(effective, numeric, EFFECT_KEYS);
  if (linked.effectNoStyle !== true && hasPresent(linked, EFFECT_KEYS)) {
    // An authored empty list and an unsupported concrete recipe both clear a
    // numeric effect. Only linked effectRef idx=0 (`effectNoStyle`) falls
    // through. This is deliberately a whole-component replacement: unlike a
    // fill or line, CT_EffectList is an ordered composite recipe rather than a
    // bag of independently inherited effect children.
    for (const key of EFFECT_KEYS) delete (effective as Record<string, unknown>)[key];
    copyPresent(effective, linked, EFFECT_KEYS);
  }

  // CT_StyleEntry modifiers belong to the linked role itself, not to any one
  // fill/line/effect atom. Preserve them independently so the renderer can
  // decide whether a local spPr is allowed to replace an absent linked atom.
  copyPresent(effective, linked, MODIFIER_KEYS);

  return effective;
}

/** Materialize the renderer-facing role table without losing source layers. */
export function withEffectiveChartStyleRoles(chart: ChartModel): ChartModel {
  const numeric = chart.classicChartStyleRoles;
  const linked = chart.linkedChartStyleRoles ?? chart.chartStyleRoles;
  if (!numeric) return chart;

  const roleNames = new Set<ChartStyleRole>([
    ...Object.keys(numeric) as ChartStyleRole[],
    ...Object.keys(linked ?? {}) as ChartStyleRole[],
  ]);
  const roles: Partial<Record<ChartStyleRole, ChartExElementStyle>> = {};
  for (const role of roleNames) {
    const effective = effectiveChartStyleRole(numeric[role], linked?.[role]);
    if (effective) roles[role] = effective;
  }
  const pointRoles: Partial<Record<ChartStyleRole, ChartExElementStyle>> = {};
  for (const role of ['dataPoint', 'dataPoint3D', 'dataPointLine', 'dataPointMarker'] as const) {
    const effective = effectiveChartStyleRole(
      chart.classicVaryingPointChartStyleRoles?.[role] ?? numeric[role],
      linked?.[role],
    );
    if (effective) pointRoles[role] = effective;
  }
  const pointRolesByGroup = chart.classicVaryingPointChartStyleRolesByGroup?.map(
    (groupRoles) => {
      if (!groupRoles) return null;
      const effectiveGroup: Partial<Record<ChartStyleRole, ChartExElementStyle>> = {};
      for (const role of ['dataPoint', 'dataPoint3D', 'dataPointLine', 'dataPointMarker'] as const) {
        const effective = effectiveChartStyleRole(
          groupRoles[role], linked?.[role],
        );
        if (effective) effectiveGroup[role] = effective;
      }
      return effectiveGroup;
    },
  );

  // Classic families share the same mark painters as ChartEx. These aliases
  // are renderer adapters only; the two source layers remain separately
  // available on the model for precedence and diagnostics.
  return {
    ...chart,
    linkedChartStyleRoles: linked,
    chartStyleRoles: roles,
    varyingPointChartStyleRoles: Object.keys(pointRoles).length > 0 ? pointRoles : undefined,
    varyingPointChartStyleRolesByGroup: pointRolesByGroup,
    // A classic chart may already carry the raw linked role in these legacy
    // ChartEx-named adapter fields.  Once a numeric classic style exists the
    // adapters must point at the linked-over-numeric result, otherwise a raw
    // partial/NoStyle role bypasses the numeric fallback in shared painters.
    // True ChartEx models have no classic role table and returned above.
    chartexDataPointStyle: roles.dataPoint,
    chartexDataPointLineStyle: roles.dataPointLine,
    chartexSeriesLineStyle: roles.seriesLine,
    chartexDataPointMarkerStyle: roles.dataPointMarker,
  };
}

/**
 * Return the raw Office 2013+ Chart Style role which owns CT_StyleEntry
 * modifiers.  The renderer-facing role table may already be linked-over-
 * numeric, so it cannot answer whether an authored noFill/no-line is allowed
 * to replace the linked role.  Numeric-only classic styles deliberately
 * return undefined here: their paint remains directly replaceable.
 */
export function rawLinkedChartStyleRole(
  chart: ChartModel,
  role: ChartStyleRole,
): ChartExElementStyle | undefined {
  return chart.linkedChartStyleRoles?.[role]
    ?? (chart.classicChartStyleRoles == null
      ? chart.chartStyleRoles?.[role]
      : undefined);
}

/** Whether the classic plot group that owns one flattened series uses the
 * point formatting-index domain from ECMA-376 §21.2.3.46 Table 6. Keep this
 * group-local: a varying pie/bar group in a combo chart must not recolour an
 * unrelated series owned by another group. */
export function chartSeriesVariesByPoint(chart: ChartModel, seriesIndex: number): boolean {
  const group = chartPlotGroupForSeries(chart, seriesIndex);
  if (group) {
    if (group.kind === 'pie' || group.kind === 'pie3D'
      || group.kind === 'doughnut' || group.kind === 'ofPie') {
      return group.varyColors !== false;
    }
    if (group.kind === 'bubble') {
      return group.seriesCount === 1 && group.varyColors !== false;
    }
    if (group.kind === 'line' || group.kind === 'scatter') {
      return group.seriesCount === 1 && group.varyColors === true;
    }
    if (group.kind === 'radar') {
      // Excel uses point formatting for lone standard/marker radar series,
      // but a filled radar remains one series-owned polygon even when
      // c:varyColors=1. This was checked with Office-produced true/false
      // counterexamples; applying point roles to the fill would invent a
      // meaning the classic chart model cannot paint.
      return group.radarStyle !== 'filled'
        && group.seriesCount === 1 && group.varyColors === true;
    }
    return (group.kind === 'bar' || group.kind === 'bar3D')
      && group.seriesCount === 1 && group.varyColors === true;
  }
  // Backward compatibility for hand-built ChartModel values that predate
  // plotGroups. Parsed OOXML always takes the bounded ownership path above.
  return chart.chartType === 'pie' || chart.chartType === 'pie3D'
    || chart.chartType === 'doughnut' || chart.chartType === 'ofPie'
    ? chart.varyColors !== false
    : chart.chartType === 'bubble'
      ? chart.series.length === 1 && chart.varyColors !== false
    : chart.varyColors === true && chart.series.length === 1
      && (chart.chartType.includes('Bar')
        || chart.chartType === 'line' || chart.chartType === 'stackedLine'
        || chart.chartType === 'stackedLinePct' || chart.chartType === 'scatter'
        || (chart.chartType === 'radar' && chart.radarStyle !== 'filled'));
}

/** Select the effective data-point role in its owning formatting-index domain. */
export function chartDataPointStyleRole(
  chart: ChartModel,
  role: 'dataPoint' | 'dataPoint3D' | 'dataPointLine' | 'dataPointMarker',
  seriesIndex: number,
): ChartExElementStyle | undefined {
  if (!chartSeriesVariesByPoint(chart, seriesIndex)) return chart.chartStyleRoles?.[role];
  const groupIndex = chartStyleGroupIndex(chart, seriesIndex);
  const groupRoles = groupIndex >= 0
    ? chart.varyingPointChartStyleRolesByGroup?.[groupIndex]
    : undefined;
  if (groupRoles != null) return groupRoles[role];
  return chart.varyingPointChartStyleRoles?.[role] ?? chart.chartStyleRoles?.[role];
}
