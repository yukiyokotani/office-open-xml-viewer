// The swap seam. The low-level UAX#9 surface is isolated behind this interface
// so a from-scratch implementation can be replaced by a thin third-party
// adapter (e.g. bidi-js: getEmbeddingLevels / getReorderSegments /
// getMirroredCharacter) without touching renderers or `toVisualSegments`,
// which sit ABOVE this seam.

import type { BaseDirection, BidiClass, BidiLevels } from './types.js';
import { createUax9Engine } from './uax9/index.js';

export interface BidiEngine {
  /**
   * UAX#9 P2-P3 + X..I rules. Returns the resolved embedding level per code
   * unit and the paragraph base level. Code units removed by rule X9 MUST be
   * marked with {@link REMOVED_LEVEL} (255). Mirrors bidi-js
   * `getEmbeddingLevels`.
   *
   * `classOverride` is an optional UAX#9 §4.3 HL1 higher-level-protocol hook:
   * a per-CODE-UNIT array (indexed identically to `text`; surrogate-pair halves
   * share one entry on the leading unit, trailing unit ignored) whose non-null
   * entries replace the assigned Bidi_Class of that code point before the
   * algorithm runs. Used e.g. to classify European digits as AN in Arabic
   * complex-script context (Word behaviour). Omit it to run pure UAX#9.
   */
  computeLevels(
    text: string,
    base: BaseDirection,
    classOverride?: ReadonlyArray<BidiClass | null | undefined>,
  ): { levels: BidiLevels; paragraphLevel: number };

  /**
   * UAX#9 L2. Given resolved levels for the line range [start, end), returns the
   * visual-order permutation: result[i] is the logical index drawn at visual
   * position i. Mirrors bidi-js `getReorderSegments` / `getReorderedIndices`.
   */
  reorderVisual(levels: BidiLevels, start: number, end: number): number[];

  /**
   * UAX#9 L4. Mirror glyph for a character displayed at an odd (RTL) level, or
   * null when the character has no mirror. Mirrors bidi-js
   * `getMirroredCharacter`.
   */
  getMirror(codePoint: number): number | null;
}

let current: BidiEngine | null = null;
let lifecycleToken: object = Object.freeze({});

/** Internal invalidation identity, not an embedding-level/result cache.
 * Providers must be deterministic for equal explicit inputs during a
 * synchronous layout. After changing captured semantic configuration, call
 * setBidiEngine again, even when retaining the same adapter object. Hidden
 * closure mutation without a lifecycle call cannot be detected by consumers. */
export function getBidiEngineLifecycleToken(): object {
  return lifecycleToken;
}

/** The active engine. Lazily instantiates the default UAX#9 engine. */
export function getDefaultBidiEngine(): BidiEngine {
  if (current === null) current = createUax9Engine();
  return current;
}

/** Swap the active engine (tests, or a third-party adapter). Call again after
 * changing captured semantic configuration, even with the same object, so
 * synchronous layout consumers invalidate derived ordering. */
export function setBidiEngine(engine: BidiEngine): void {
  current = engine;
  lifecycleToken = Object.freeze({});
}

/** Restore the built-in UAX#9 engine. */
export function resetBidiEngine(): void {
  current = null;
  lifecycleToken = Object.freeze({});
}
