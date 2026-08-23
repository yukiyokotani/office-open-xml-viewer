import { MAX_CANVAS_CHART_POINTS } from './resource-limits.js';

// The package parser already applies its XML-depth ceiling. Keep the same kind
// of stack-safety boundary for caller-constructed public ChartModel objects.
export const MAX_CANVAS_HIERARCHY_DEPTH = 512;

/** A node in the ChartEx hierarchy tree. `layoutWeight` is the overflow-safe
 * aggregate used for treemap areas and sunburst angles; `value` is the finite,
 * saturating aggregate exposed to data-label formatting. */
export interface SunburstNode {
  label: string;
  layoutWeight: number;
  value: number;
  depth: number;
  children: SunburstNode[];
  branchIndex: number;
  labelIndex: number;
  a0: number;
  a1: number;
}

/** Preflight flat hierarchy input before allocating the interned tree. */
export function hierarchyInputTooLarge(rows: readonly { path: readonly string[] }[]): boolean {
  if (rows.length > MAX_CANVAS_CHART_POINTS) return true;
  let segments = 0;
  for (const row of rows) {
    if (row.path.length > MAX_CANVAS_HIERARCHY_DEPTH) return true;
    if (segments > MAX_CANVAS_CHART_POINTS - row.path.length) return true;
    segments += row.path.length;
  }
  return false;
}

/** Fold flat `path`/`size` rows into a source-ordered hierarchy tree. */
export function buildSunburstTree(
  rows: { path: string[]; size: number }[],
  preserveTerminalDuplicates = false,
): SunburstNode {
  const root: SunburstNode = {
    label: '', layoutWeight: 0, value: 0, depth: -1, children: [], branchIndex: -1,
    labelIndex: -1, a0: 0, a1: 0,
  };
  const maxRowValue = rows.reduce((max, row) => (
    Number.isFinite(row.size) && row.size > max ? row.size : max
  ), 0);
  const safeSum = (left: number, right: number): number => (
    left > Number.MAX_VALUE - right ? Number.MAX_VALUE : left + right
  );
  const childIndexes = new WeakMap<SunburstNode, Map<string, SunburstNode>>();
  for (const row of rows) {
    const rowValue = Number.isFinite(row.size) && row.size > 0 ? row.size : 0;
    const rowWeight = maxRowValue > 0 ? rowValue / maxRowValue : 0;
    let node = root;
    for (let depth = 0; depth < row.path.length; depth++) {
      const label = row.path[depth];
      let index = childIndexes.get(node);
      if (!index) {
        index = new Map();
        childIndexes.set(node, index);
      }
      const preserveNode = preserveTerminalDuplicates && depth === row.path.length - 1;
      let child = preserveNode ? undefined : index.get(label);
      if (!child) {
        child = {
          label,
          layoutWeight: 0,
          value: 0,
          depth,
          children: [],
          branchIndex: depth === 0 ? node.children.length : node.branchIndex,
          labelIndex: -1,
          a0: 0,
          a1: 0,
        };
        node.children.push(child);
        if (!preserveNode) index.set(label, child);
      }
      child.layoutWeight += rowWeight;
      child.value = safeSum(child.value, rowValue);
      node = child;
    }
  }
  root.layoutWeight = root.children.reduce((sum, child) => sum + child.layoutWeight, 0);
  root.value = root.children.reduce((sum, child) => safeSum(sum, child.value), 0);
  let nextLabelIndex = 0;
  const pending = [...root.children].reverse();
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    node.labelIndex = nextLabelIndex++;
    for (let index = node.children.length - 1; index >= 0; index--) {
      pending.push(node.children[index]);
    }
  }
  return root;
}

/** Assign each node's angular span from its children's relative weights. */
export function layoutSunburstAngles(root: SunburstNode): void {
  const pending = [root];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    let total = 0;
    for (const child of node.children) total += child.layoutWeight;
    if (total <= 0) continue;
    let angle = node.a0;
    for (const child of node.children) {
      const sweep = ((node.a1 - node.a0) * child.layoutWeight) / total;
      child.a0 = angle;
      child.a1 = angle + sweep;
      angle = child.a1;
      pending.push(child);
    }
  }
}

/** Maximum ring depth below the synthetic root. */
export function sunburstMaxDepth(root: SunburstNode): number {
  let maxDepth = root.depth;
  const pending = [root];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    maxDepth = Math.max(maxDepth, node.depth);
    for (let index = node.children.length - 1; index >= 0; index--) {
      pending.push(node.children[index]);
    }
  }
  return maxDepth;
}
