import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from '../document-layout.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { createLayoutServices } from '../layout-runtime.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from '../testing/synthetic-document.js';
import { layoutDocumentInputAsync } from './document.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { PaginationAbortError } from './pagination-scheduler.js';
import { setDocumentLayoutValidation } from './validation-policy.js';
import type { DocumentLayout } from './types.js';

// ─────────────────────────────────────────────────────────────────────────────
// Time-sliced pagination must be a scheduling change and nothing else. The
// generator decides where suspension is safe (between body entries); the
// scheduler only decides whether to take those opportunities. So for ANY
// policy — including the adversarial "yield at every single opportunity" —
// the layout has to be byte-identical to the synchronous one.
//
// These fixtures each drive a different convergence solver, so the equivalence
// is checked across single-pass, reserve-converged and field-converged
// pagination rather than only the simple case.
// ─────────────────────────────────────────────────────────────────────────────

const CURRENT_DATE_MS = 1_700_000_000_000;

const SHAPES: readonly (readonly [SyntheticDocumentShape, number])[] = [
  ['plain', 30],
  ['header-footer', 30],
  ['fields', 30],
  ['tables', 8],
  ['long-paragraphs', 2],
];

/** Lay out asynchronously with a policy that suspends at every opportunity. */
async function layoutAdversarially(
  shape: SyntheticDocumentShape,
  paragraphs: number,
  extra: { signal?: AbortSignal; onProgress?: (pages: number) => void } = {},
): Promise<DocumentLayout> {
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs }));
  const services = createLayoutServices(source);
  return layoutDocumentInputAsync(
    source.bodyLayoutInput,
    services,
    normalizeLayoutOptions(undefined, CURRENT_DATE_MS),
    {
      // now() always reports a spent budget, so every suspension point yields.
      now: () => Number.MAX_SAFE_INTEGER,
      sliceMs: 0,
      yieldToHost: () => Promise.resolve(),
      ...extra,
    },
  );
}

function layoutSynchronously(
  shape: SyntheticDocumentShape,
  paragraphs: number,
): DocumentLayout {
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs }));
  const services = createLayoutServices(source);
  return layoutDocument(
    source,
    services,
    normalizeLayoutOptions(undefined, CURRENT_DATE_MS),
  ) as DocumentLayout;
}

beforeAll(() => {
  installStubCanvas();
});

afterAll(() => {
  setDocumentLayoutValidation(true);
});

describe('time-sliced pagination equals blocking pagination', () => {
  for (const [shape, paragraphs] of SHAPES) {
    it(`${shape}`, async () => {
      const blocking = layoutSynchronously(shape, paragraphs);
      const sliced = await layoutAdversarially(shape, paragraphs);
      expect(sliced.pages.length).toBe(blocking.pages.length);
      expect(layoutFingerprint(sliced)).toBe(layoutFingerprint(blocking));
    }, 300_000);
  }
});

describe('pagination scheduler', () => {
  it('reports committed pages while laying out', async () => {
    const seen: number[] = [];
    const layout = await layoutAdversarially('plain', 30, {
      onProgress: (pages) => { seen.push(pages); },
    });
    expect(seen.length).toBeGreaterThan(0);
    expect(Math.max(...seen)).toBeGreaterThan(0);
    // Progress is a count of committed pages, so it can never exceed the pages
    // the finished layout actually has.
    expect(Math.max(...seen)).toBeLessThanOrEqual(layout.pages.length);
  }, 300_000);

  it('reports progress monotonically within a single pass', async () => {
    // `plain` has no pagination fields, page-owned anchors or continuous
    // sections, so it converges in one pass and progress cannot restart.
    const seen: number[] = [];
    await layoutAdversarially('plain', 30, {
      onProgress: (pages) => { seen.push(pages); },
    });
    const decreases = seen.filter((value, index) => index > 0 && value < seen[index - 1]!);
    expect(decreases).toEqual([]);
  }, 300_000);

  it('aborts at the next suspension point', async () => {
    const controller = new AbortController();
    let steps = 0;
    await expect(layoutAdversarially('plain', 60, {
      signal: controller.signal,
      onProgress: () => {
        steps += 1;
        if (steps === 3) controller.abort();
      },
    })).rejects.toBeInstanceOf(PaginationAbortError);
  }, 300_000);

  it('rejects rather than resolving a partial layout when aborted', async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(layoutAdversarially('plain', 10, { signal: controller.signal }))
      .rejects.toBeInstanceOf(PaginationAbortError);
  }, 300_000);
});
