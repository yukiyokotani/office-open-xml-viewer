interface SparseStyleIndexCache {
  compact: WeakMap<readonly number[], ReadonlyMap<number, number>>;
  membership: WeakMap<readonly number[], ReadonlySet<number>>;
}

let activeCache: SparseStyleIndexCache | undefined;

/** Scope memoized sparse-index materialization to one synchronous render.
 * Public ChartModel/style arrays remain mutable between renders, so retaining
 * these maps globally would make an in-place caller update observe stale style
 * ownership. Nested renderer helpers reuse the already-active scope. */
export function withSparseStyleIndexCache<T>(run: () => T): T {
  if (activeCache) return run();
  activeCache = { compact: new WeakMap(), membership: new WeakMap() };
  try {
    return run();
  } finally {
    activeCache = undefined;
  }
}

/** Map a sparse source formatting index to its compact materialized slot. */
export function compactStyleIndex(
  indices: readonly number[],
  sourceIndex: number,
): number {
  if (!activeCache) return indices.indexOf(sourceIndex);
  let lookup = activeCache.compact.get(indices);
  if (!lookup) {
    const materialized = new Map<number, number>();
    for (let index = 0; index < indices.length; index++) {
      const value = indices[index];
      if (!materialized.has(value)) materialized.set(value, index);
    }
    lookup = materialized;
    activeCache.compact.set(indices, lookup);
  }
  return lookup.get(sourceIndex) ?? -1;
}

/** Bounded O(1) membership for repeated semantic-fallback decisions. */
export function styleIndexSetHas(
  indices: readonly number[] | null | undefined,
  index: number,
): boolean {
  if (!indices) return false;
  if (!activeCache) return indices.includes(index);
  let lookup = activeCache.membership.get(indices);
  if (!lookup) {
    lookup = new Set(indices);
    activeCache.membership.set(indices, lookup);
  }
  return lookup.has(index);
}
