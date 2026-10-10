import { NativeReadingSceneError } from './native-reading-block-scene.js';

/** Reading blocks may move whole, but cannot use ordinary Word overflow
 * allowances. Notes and narrowed page bands are part of the actual destination
 * capacity. A fresh-column or reserve failure revokes the complete candidate. */
export function nativeReadingRequiresFreshColumn(requiredPt: number, availablePt: number, freshPt: number, canRelocate: boolean, reserveFits: boolean): boolean {
  if (![requiredPt, availablePt, freshPt].every(value => Number.isFinite(value) && value >= 0))
    throw new NativeReadingSceneError('placement', 'reading block capacity must be finite and nonnegative');
  if (requiredPt <= availablePt && reserveFits) return false;
  if (canRelocate && requiredPt <= freshPt) return true;
  throw new NativeReadingSceneError('placement', 'complete reading block and its notes cannot fit the destination column');
}
