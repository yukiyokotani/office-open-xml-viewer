import { describe, expect, it } from 'vitest';
import {
  addUprightTableBox,
  emptyUprightBlockPageIndex,
  type UprightTableBoxVersions,
} from './upright-block-index.js';

describe('upright block page index', () => {
  const box = (at: number) => ({ xPt: 60 + 10 * at, yPt: 18 * at, widthPt: 30, heightPt: 180 });
  const filled = (versions: UprightTableBoxVersions, count: number, shift = 0) => {
    const index = emptyUprightBlockPageIndex();
    for (let at = 0; at < count; at += 1) {
      addUprightTableBox(index, at, { ...box(at), xPt: box(at).xPt + (at === count - 1 ? shift : 0) }, versions);
    }
    return index;
  };

  it('versions the history of table boxes by content, in a key that does not grow with it', () => {
    // What keys a drawing-anchoring paragraph's acquisition: the same boxes
    // give the same version in another layout of the page, any other history
    // a different one, and the key stays a number however many boxes it
    // stands for.
    const versions: UprightTableBoxVersions = new Map();
    for (const count of [4, 16, 256]) {
      const first = filled(versions, count);
      const again = filled(versions, count);
      expect(again.tableBoxesVersion).toBe(first.tableBoxesVersion);
      expect(filled(versions, count, 1).tableBoxesVersion).not.toBe(first.tableBoxesVersion);
      expect(filled(versions, count - 1).tableBoxesVersion).not.toBe(first.tableBoxesVersion);
      expect(JSON.stringify(first.tableBoxesVersion).length).toBeLessThanOrEqual(6);
    }
    expect(emptyUprightBlockPageIndex().tableBoxesVersion).toBe(0);
  });
});
