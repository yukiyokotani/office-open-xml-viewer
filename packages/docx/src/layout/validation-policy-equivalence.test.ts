import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from '../document-layout.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from '../testing/synthetic-document.js';
import { layoutFingerprint } from './invariants.js';
import { setDocumentLayoutValidation } from './validation-policy.js';
import type { DocumentLayout } from './types.js';

// ─────────────────────────────────────────────────────────────────────────────
// The load-bearing invariant behind making the retained-layout contract checks
// development-only: turning them off must change PERFORMANCE ONLY. The produced
// layout has to stay byte-identical, because the variant store caches it, the
// renderer paints from it, and the VRT corpus is pinned to it.
//
// `layoutFingerprint` is the canonical comparison (numbers rounded to 6dp,
// object keys sorted, diagnostic prose excluded), so this is a much stronger
// check than spot-asserting page counts. Each synthetic shape drives a
// different one of `paginateBody`'s convergence solvers — see
// `testing/synthetic-document.ts`.
// ─────────────────────────────────────────────────────────────────────────────

const SHAPES: readonly (readonly [SyntheticDocumentShape, number])[] = [
  ['plain', 40],
  ['header-footer', 40],
  ['fields', 40],
  ['tables', 12],
  ['long-paragraphs', 3],
];

beforeAll(() => {
  installStubCanvas();
});

afterAll(() => {
  setDocumentLayoutValidation(true);
});

describe('layout is identical with and without validation', () => {
  for (const [shape, paragraphs] of SHAPES) {
    it(`${shape}`, () => {
      // Two independent models: a shared one could let the first run's freezing
      // mask a difference in the second.
      setDocumentLayoutValidation(true);
      const validated = layoutDocument(syntheticDocxModel(shape, { paragraphs }));

      setDocumentLayoutValidation(false);
      const unvalidated = layoutDocument(syntheticDocxModel(shape, { paragraphs }));

      setDocumentLayoutValidation(true);
      expect(unvalidated.pages.length).toBe(validated.pages.length);
      expect(layoutFingerprint(unvalidated as DocumentLayout))
        .toBe(layoutFingerprint(validated as DocumentLayout));
    }, 300_000);
  }

  it('freezes the retained graph either way', () => {
    setDocumentLayoutValidation(false);
    const layout = layoutDocument(syntheticDocxModel('plain', { paragraphs: 8 }));
    setDocumentLayoutValidation(true);
    expect(Object.isFrozen(layout)).toBe(true);
    expect(Object.isFrozen(layout.pages)).toBe(true);
    expect(Object.isFrozen(layout.pages[0])).toBe(true);
    expect(Object.isFrozen(layout.pages[0].geometry)).toBe(true);
    expect(Object.isFrozen(layout.pages[0].layers.body)).toBe(true);
  });
});
