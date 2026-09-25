import type { ChartModel } from '../types/chart.js';
import {
  buildSunburstTree,
  hierarchyInputTooLarge,
  layoutSunburstAngles,
  type SunburstNode,
} from './chart-ex-hierarchy.js';
import { resolveChartExLabel, type ResolvedChartExLabel } from './chart-ex-label.js';

export interface ChartExHierarchyLabelSite {
  label: ResolvedChartExLabel;
  styleIndex: number;
  linkedStyleIndex: number;
}

export type ChartExHierarchyLabelVisitResult = 'not-hierarchy' | 'too-large' | 'ok';

export interface ChartExHierarchyBodySite {
  node: SunburstNode;
  kind: 'sunburst' | 'treemap';
  paintsBody: boolean;
}

/** Enumerate hierarchy geometry sites without allocating a second flattened
 * node array. `paintsBody` matches the current ChartEx family semantics:
 * sunburst paints every visible wedge, while treemap paints leaves and banner
 * parent bands but not overlapping parent caption containers. */
export function visitChartExHierarchyBodySites(
  chart: ChartModel,
  visit: (site: ChartExHierarchyBodySite) => void,
): ChartExHierarchyLabelVisitResult {
  const hierarchy = chart.chartexSunburst
    ? { rows: chart.chartexSunburst.rows, kind: 'sunburst' as const }
    : chart.chartexTreemap
      ? { rows: chart.chartexTreemap.rows, kind: 'treemap' as const }
      : undefined;
  if (!hierarchy) return 'not-hierarchy';
  if (hierarchy.rows.length === 0) return 'ok';
  if (hierarchyInputTooLarge(hierarchy.rows)) return 'too-large';
  const root = buildSunburstTree(hierarchy.rows, hierarchy.kind === 'treemap');
  if (root.layoutWeight <= 0 || root.children.length === 0) return 'ok';
  if (hierarchy.kind === 'sunburst') {
    root.a0 = -Math.PI / 2;
    root.a1 = root.a0 + Math.PI * 2;
    layoutSunburstAngles(root);
  }
  const parentMode = chart.chartexTreemap?.parentLabelLayout ?? 'overlapping';
  const series = chart.series[0];
  const overrides = new Map((series?.dataLabelOverrides ?? []).map(item => [item.idx, item]));
  const pending = [...root.children];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    for (const child of node.children) pending.push(child);
    if (node.layoutWeight <= 0
      || (hierarchy.kind === 'sunburst' && node.a1 - node.a0 <= 1e-4)) continue;
    const bannerLabel = hierarchy.kind === 'treemap'
      && node.children.length > 0
      && parentMode === 'banner'
      ? resolveChartExLabel(
          chart, series, node.labelIndex, node.label, node.value,
          { visible: true, showVal: false, showCatName: true }, overrides, true,
        )
      : null;
    visit({
      node,
      kind: hierarchy.kind,
      paintsBody: hierarchy.kind === 'sunburst'
        || node.children.length === 0
        || bannerLabel != null,
    });
  }
  return 'ok';
}

/** Enumerate the exact interned hierarchy nodes that can reach the shared
 * label painter. Resource preflight and picture prefetch use this same plan so
 * a deep single source row cannot multiply label work outside their budgets. */
export function visitChartExHierarchyLabelSites(
  chart: ChartModel,
  visit: (site: ChartExHierarchyLabelSite) => void,
): ChartExHierarchyLabelVisitResult {
  const hierarchy = chart.chartexSunburst
    ? { rows: chart.chartexSunburst.rows, kind: 'sunburst' as const }
    : chart.chartexTreemap
      ? { rows: chart.chartexTreemap.rows, kind: 'treemap' as const }
      : undefined;
  if (!hierarchy) return 'not-hierarchy';
  if (hierarchy.rows.length === 0) return 'ok';
  if (hierarchyInputTooLarge(hierarchy.rows)) return 'too-large';

  const root = buildSunburstTree(hierarchy.rows, hierarchy.kind === 'treemap');
  if (root.layoutWeight <= 0 || root.children.length === 0) return 'ok';
  if (hierarchy.kind === 'sunburst') {
    root.a0 = -Math.PI / 2;
    root.a1 = root.a0 + Math.PI * 2;
    layoutSunburstAngles(root);
  }
  const series = chart.series[0];
  const overrides = new Map((series?.dataLabelOverrides ?? []).map(item => [item.idx, item]));
  const parentMode = chart.chartexTreemap?.parentLabelLayout ?? 'overlapping';
  const pending = [...root.children];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    for (const child of node.children) pending.push(child);
    if (node.layoutWeight <= 0
      || (hierarchy.kind === 'sunburst' && node.a1 - node.a0 <= 1e-4)) continue;

    let label;
    if (hierarchy.kind === 'sunburst') {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: false, showVal: false, showCatName: false }, overrides,
      );
    } else if (node.children.length > 0) {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: parentMode !== 'none', showVal: false, showCatName: true }, overrides,
      );
      if (parentMode === 'overlapping' && node.depth !== 0) label = null;
    } else {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: false, showVal: false, showCatName: false }, overrides,
      );
    }
    if (label) visit({
      label,
      styleIndex: node.labelIndex,
      linkedStyleIndex: overrides.has(node.labelIndex)
        ? node.labelIndex : series?.chartexFormatIdx ?? 0,
    });
  }
  return 'ok';
}
