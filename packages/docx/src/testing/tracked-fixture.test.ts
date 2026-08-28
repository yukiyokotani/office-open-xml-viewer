import { beforeAll, describe, expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { paginateBody } from '../layout/body-paginator.js';
import { layoutFingerprint } from '../layout/invariants.js';
import { normalizeLayoutOptions } from '../layout/options.js';
import { installStubCanvas, syntheticDocxModel } from './synthetic-document.js';
import type { DocumentLayout } from '../layout/types.js';

// ─────────────────────────────────────────────────────────────────────────────
// The `tracked` fixtures exist to distinguish the two `showTrackedChanges`
// layout variants, so their usefulness is itself worth pinning: if hiding
// deletions ever stopped changing pagination for this content, every test built
// on the fixture would keep passing while testing nothing.
// ─────────────────────────────────────────────────────────────────────────────

const CURRENT_DATE_MS = 1_700_000_000_000;

function layoutOf(shape: 'tracked' | 'tracked-fields', markup: boolean): DocumentLayout {
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs: 120 }));
  return paginateBody(
    source.bodyLayoutInput,
    createLayoutServices(source),
    normalizeLayoutOptions(undefined, CURRENT_DATE_MS, markup),
  );
}

beforeAll(() => {
  installStubCanvas();
});

describe('tracked-changes fixture', () => {
  it('paginates differently in the final and markup views', () => {
    const finalView = layoutOf('tracked', false);
    const markupView = layoutOf('tracked', true);
    // The markup view keeps deleted text, so it must need more pages. An equal
    // count would mean the fixture cannot tell the variants apart.
    expect(markupView.pages.length).toBeGreaterThan(finalView.pages.length);
    expect(layoutFingerprint(markupView)).not.toBe(layoutFingerprint(finalView));
  }, 300_000);

  it('is deterministic per variant', () => {
    expect(layoutFingerprint(layoutOf('tracked', true)))
      .toBe(layoutFingerprint(layoutOf('tracked', true)));
    expect(layoutFingerprint(layoutOf('tracked', false)))
      .toBe(layoutFingerprint(layoutOf('tracked', false)));
  }, 300_000);

  it('reports pagination fields only for the tracked-fields shape', () => {
    expect(layoutSourceStore(syntheticDocxModel('tracked', { paragraphs: 8 }))
      .hasPaginationFields).toBe(false);
    expect(layoutSourceStore(syntheticDocxModel('tracked-fields', { paragraphs: 8 }))
      .hasPaginationFields).toBe(true);
  });
});
