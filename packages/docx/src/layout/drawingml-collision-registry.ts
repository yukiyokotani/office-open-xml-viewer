import type {
  DrawingMLCollisionEntryPt,
  DrawingMLCollisionRegistryDeltaPt,
  DrawingMLCollisionRegistrySnapshotPt,
  FloatRegistryCoordinateSpace,
} from './types.js';

function validateEntry(entry: DrawingMLCollisionEntryPt): void {
  if (entry.occurrenceId.length === 0) {
    throw new Error('DrawingML collision occurrence ID must not be empty');
  }
  const { xPt, yPt, widthPt, heightPt } = entry.bounds;
  if (![xPt, yPt, widthPt, heightPt].every(Number.isFinite)
    || widthPt < 0
    || heightPt < 0) {
    throw new Error(`DrawingML collision bounds are invalid: ${entry.occurrenceId}`);
  }
  if (
    (entry.horizontalOwnership !== 'page' && entry.horizontalOwnership !== 'host')
    || (entry.verticalOwnership !== 'page' && entry.verticalOwnership !== 'host')
  ) {
    throw new Error(`DrawingML collision ownership is invalid: ${entry.occurrenceId}`);
  }
}

function snapshotEntry(
  entry: DrawingMLCollisionEntryPt,
): DrawingMLCollisionEntryPt {
  validateEntry(entry);
  return Object.freeze({
    occurrenceId: entry.occurrenceId,
    bounds: Object.freeze({ ...entry.bounds }),
    horizontalOwnership: entry.horizontalOwnership,
    verticalOwnership: entry.verticalOwnership,
    ...(entry.relativeHeight !== undefined
      ? { relativeHeight: entry.relativeHeight }
      : {}),
  });
}

export function createDrawingMLCollisionRegistry(
  flowDomainId: string,
  coordinateSpace: FloatRegistryCoordinateSpace,
): DrawingMLCollisionRegistrySnapshotPt {
  return Object.freeze({
    coordinateSpace,
    flowDomainId,
    entries: Object.freeze([]),
  });
}

export function drawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  entries: readonly DrawingMLCollisionEntryPt[],
): DrawingMLCollisionRegistryDeltaPt {
  return Object.freeze({
    coordinateSpace: snapshot.coordinateSpace,
    flowDomainId: snapshot.flowDomainId,
    baseEntries: snapshot.entries,
    baseEntryCount: snapshot.entries.length,
    entries: Object.freeze(entries.map(snapshotEntry)),
  });
}

export function validateDrawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  delta: DrawingMLCollisionRegistryDeltaPt,
): void {
  if (delta.coordinateSpace !== snapshot.coordinateSpace) {
    throw new Error('DrawingML collision registry coordinate space mismatch');
  }
  if (delta.flowDomainId !== snapshot.flowDomainId) {
    throw new Error('DrawingML collision registry flow domain mismatch');
  }
  if (
    delta.baseEntries !== snapshot.entries
    || delta.baseEntryCount !== snapshot.entries.length
  ) {
    throw new Error('DrawingML collision registry delta is stale');
  }
  const known = new Set(snapshot.entries.map((entry) => entry.occurrenceId));
  for (const entry of delta.entries) {
    validateEntry(entry);
    if (known.has(entry.occurrenceId)) {
      throw new Error(`DrawingML collision occurrence committed twice: ${entry.occurrenceId}`);
    }
    known.add(entry.occurrenceId);
  }
}

export function applyDrawingMLCollisionRegistryDelta(
  snapshot: DrawingMLCollisionRegistrySnapshotPt,
  delta: DrawingMLCollisionRegistryDeltaPt,
): DrawingMLCollisionRegistrySnapshotPt {
  validateDrawingMLCollisionRegistryDelta(snapshot, delta);
  const entries = Object.freeze([...snapshot.entries, ...delta.entries]);
  extendedFrom.set(entries, snapshot.entries);
  return Object.freeze({
    coordinateSpace: snapshot.coordinateSpace,
    flowDomainId: snapshot.flowDomainId,
    entries,
  });
}

/** The entries a registry's entries were extended from: a prefix of them. */
const extendedFrom = new WeakMap<readonly DrawingMLCollisionEntryPt[], readonly DrawingMLCollisionEntryPt[]>();

type EntriesVersions = {
  /** Number of each (version, next entry) step: one per distinct content. */
  readonly steps: Map<string, number>;
  readonly known: WeakMap<readonly DrawingMLCollisionEntryPt[], number>;
};
const versionsByScope = new WeakMap<object, EntriesVersions>();

/**
 * A number that stands for the exact content of `entries` within `scope`
 * (one acquisition cache): two lists get the same number exactly when they
 * hold equal entries in the same order, however they were put together.
 * Zero is the empty list.
 *
 * The number of a list is the number of the list before its last entry,
 * extended by that entry. A registry that grew from another
 * (applyDrawingMLCollisionRegistryDelta) is numbered from the one it grew
 * from, so numbering each registry of a page as it grows reads every entry
 * once.
 */
export function drawingMLCollisionEntriesVersion(
  scope: object,
  entries: readonly DrawingMLCollisionEntryPt[],
): number {
  if (entries.length === 0) return 0;
  let versions = versionsByScope.get(scope);
  if (!versions) versionsByScope.set(scope, versions = { steps: new Map(), known: new WeakMap() });
  // The lists this one grew from, back to the first that is numbered.
  const pending: Array<readonly DrawingMLCollisionEntryPt[]> = [];
  let version = 0;
  let numbered = 0;
  for (let list: readonly DrawingMLCollisionEntryPt[] | undefined = entries; list; list = extendedFrom.get(list)) {
    const known = versions.known.get(list);
    if (known !== undefined) {
      version = known;
      numbered = list.length;
      break;
    }
    pending.push(list);
  }
  for (let at = pending.length - 1; at >= 0; at -= 1) {
    const list = pending[at]!;
    for (; numbered < list.length; numbered += 1) {
      const step = `${version}|${JSON.stringify(list[numbered])}`;
      let next = versions.steps.get(step);
      if (next === undefined) versions.steps.set(step, next = versions.steps.size + 1);
      version = next;
    }
    versions.known.set(list, version);
  }
  return version;
}
