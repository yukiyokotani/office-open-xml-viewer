import { describe, it } from 'vitest';
import assert from 'node:assert/strict';
import { nativeReadingRequiresFreshColumn } from './native-reading-admission.js';
describe('complete reading block pagination', () => {
  it('moves a complete block when the remaining column is short and preserves exact fit', () => {
    assert.equal(nativeReadingRequiresFreshColumn(40, 30, 60, true, true), true);
    assert.equal(nativeReadingRequiresFreshColumn(40, 40, 60, true, true), false);
  });
  it('rejects oversized or immovable blocks and note-reserve overflow rather than using ordinary overflow admission', () => {
    assert.throws(() => nativeReadingRequiresFreshColumn(61, 30, 60, true, true), /cannot fit/);
    assert.throws(() => nativeReadingRequiresFreshColumn(40, 30, 60, false, true), /cannot fit/);
    assert.throws(() => nativeReadingRequiresFreshColumn(40, 60, 60, false, false), /cannot fit/);
    assert.equal(nativeReadingRequiresFreshColumn(40, 60, 60, true, false), true);
  });
});
