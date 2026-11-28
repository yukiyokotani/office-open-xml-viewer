import { describe, expect, it } from 'vitest';
import {
  applyDrawingMLCollisionRegistryDelta,
  createDrawingMLCollisionRegistry,
  drawingMLCollisionEntriesVersion,
  drawingMLCollisionRegistryDelta,
} from './drawingml-collision-registry.js';
import type { DrawingMLCollisionEntryPt, DrawingMLCollisionRegistrySnapshotPt } from './types.js';

const entry = (id: number, xPt = id * 10): DrawingMLCollisionEntryPt => ({
  occurrenceId: `drawing:${id}`,
  bounds: { xPt, yPt: 0, widthPt: 5, heightPt: 5 },
  horizontalOwnership: 'page',
  verticalOwnership: 'host',
});
const grown = (...steps: DrawingMLCollisionEntryPt[][]): DrawingMLCollisionRegistrySnapshotPt => steps.reduce(
  (registry, added) => applyDrawingMLCollisionRegistryDelta(registry, drawingMLCollisionRegistryDelta(registry, added)),
  createDrawingMLCollisionRegistry('body', 'logical-page-points'),
);

describe('number standing for the content of a drawing collision list', () => {
  it('is the same for equal entries however the list was put together', () => {
    const scope = {};
    const oneByOne = grown([entry(1)], [entry(2)], [entry(3)]);
    const inTwoSteps = grown([entry(1), entry(2)], [entry(3)]);
    const plain = [entry(1), entry(2), entry(3)];
    const version = drawingMLCollisionEntriesVersion(scope, oneByOne.entries);
    expect(version).toBeGreaterThan(0);
    expect(drawingMLCollisionEntriesVersion(scope, inTwoSteps.entries)).toBe(version);
    expect(drawingMLCollisionEntriesVersion(scope, plain)).toBe(version);
    expect(drawingMLCollisionEntriesVersion(scope, oneByOne.entries)).toBe(version);
  });

  it('differs for another entry, another order or a shorter list, and is zero for none', () => {
    const scope = {};
    const versions = [
      grown([entry(1)], [entry(2)]).entries,
      grown([entry(2)], [entry(1)]).entries,
      grown([entry(1)]).entries,
      grown([entry(1)], [entry(2, 21)]).entries,
    ].map((entries) => drawingMLCollisionEntriesVersion(scope, entries));
    expect(new Set(versions).size).toBe(versions.length);
    expect(drawingMLCollisionEntriesVersion(scope, [])).toBe(0);
  });

  it('numbers each registry of a page from the one it grew from', () => {
    // Every registry is numbered as it grows: the entries spelled out are
    // the ones added since the registry numbered before, one each.
    const scope = {};
    const stringify = JSON.stringify;
    let spelled = 0;
    JSON.stringify = ((value: unknown, ...rest: never[]) => {
      spelled += 1;
      return stringify(value, ...rest);
    }) as typeof JSON.stringify;
    try {
      let registry = createDrawingMLCollisionRegistry('body', 'logical-page-points');
      for (let id = 1; id <= 64; id += 1) {
        registry = applyDrawingMLCollisionRegistryDelta(registry, drawingMLCollisionRegistryDelta(registry, [entry(id)]));
        drawingMLCollisionEntriesVersion(scope, registry.entries);
      }
    } finally {
      JSON.stringify = stringify;
    }
    expect(spelled).toBe(64);
  });

  // Needs the collector (node --expose-gc); skipped without it.
  const collect = (globalThis as { gc?: () => void }).gc;
  it.skipIf(!collect)('does not keep the earlier registries of a page alive', async () => {
    const scope = {};
    let registry = createDrawingMLCollisionRegistry('body', 'logical-page-points');
    const earlier: WeakRef<object>[] = [];
    for (let id = 1; id <= 1000; id += 1) {
      registry = applyDrawingMLCollisionRegistryDelta(registry, drawingMLCollisionRegistryDelta(registry, [entry(id)]));
      drawingMLCollisionEntriesVersion(scope, registry.entries);
      if (id % 100 === 50) earlier.push(new WeakRef(registry.entries));
    }
    for (let round = 0; round < 2; round += 1) {
      await new Promise((resolve) => setTimeout(resolve, 10));
      collect!();
    }
    expect(earlier.filter((reference) => reference.deref() !== undefined)).toHaveLength(0);
    expect(registry.entries).toHaveLength(1000);
  });
});
