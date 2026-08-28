import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { layoutDocumentInput } from './layout/document.js';
import { normalizeLayoutOptions, type LayoutOptions } from './layout/options.js';
import { layoutDocumentProgressively } from './layout/progressive.js';
import { setDocumentLayoutValidation } from './layout/validation-policy.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from './testing/synthetic-document.js';

// ─────────────────────────────────────────────────────────────────────────────
// `showTrackedChanges` selects a different retained layout, with its own
// pagination. The bug this pins: loading a document that will be RENDERED in
// the markup view used to prime the FINAL-view layout, so the first render
// missed the cache and repaginated the whole document synchronously on the main
// thread — the progressive prefix was never selected at all, and a large
// reviewed document froze for seconds before its first page appeared.
//
// The guarantee is expressed as a builder spy: loading for a given view must
// build that view and no other. Counting builds is the only way to state it,
// since a wrong-variant build is invisible in the output — it is correct, just
// enormously expensive and thrown away.
// ─────────────────────────────────────────────────────────────────────────────

const CURRENT_DATE_MS = 1_700_000_000_000;

/** A variant store whose builder records which options it was asked for. */
function spyStore(shape: SyntheticDocumentShape, paragraphs: number) {
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs }));
  const services = createLayoutServices(source);
  const builds: LayoutOptions[] = [];
  const { store } = attachDocumentLayoutVariants({
    source,
    services,
    defaultCurrentDateMs: CURRENT_DATE_MS,
    buildLayout: (options) => {
      builds.push(options);
      return layoutDocumentInput(source.bodyLayoutInput, services, options);
    },
  });
  return { source, services, store, builds };
}

const markupOptions = normalizeLayoutOptions(undefined, CURRENT_DATE_MS, true);
const finalOptions = normalizeLayoutOptions(undefined, CURRENT_DATE_MS, false);

beforeAll(() => {
  installStubCanvas();
});

afterAll(() => {
  setDocumentLayoutValidation(true);
});

describe('progressive layout builds only the variant being viewed', () => {
  it('never builds the final view when loading for the markup view', async () => {
    const { source, services, store, builds } = spyStore('tracked', 200);

    // What `DocxDocument.load({ progressiveLayout, showTrackedChanges })` does:
    // prime the preview, then the full layout, both under the MARKUP key.
    const full = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      services,
      markupOptions,
      {
        onPreview: (preview) => { store.prime(markupOptions, preview.layout); },
      },
    );
    store.replaceIfCurrent(markupOptions, store.layoutFor(markupOptions), full);

    // Rendering a page in the markup view must hit the primed layout.
    const selected = store.select(markupOptions);
    expect(selected.layout.pages.length).toBe(full.pages.length);

    // The store's builder was never invoked: not for the final view (the bug),
    // and not for the markup view either (priming supplied it).
    expect(builds).toEqual([]);
  }, 300_000);

  it('serves the markup variant to a render after a markup-keyed prime', async () => {
    const { source, services, store, builds } = spyStore('tracked', 120);
    const markup = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      services,
      markupOptions,
    );
    store.prime(markupOptions, markup);
    builds.length = 0;

    // Selecting the OTHER view is what genuinely costs a build — and that only
    // happens when a user actually toggles.
    store.select(finalOptions);
    expect(builds).toHaveLength(1);
    expect(builds[0]!.showTrackedChanges).toBeUndefined();
    // The two variants really are different documents.
    expect(store.layoutFor(finalOptions).pages.length)
      .not.toBe(store.layoutFor(markupOptions).pages.length);
  }, 300_000);

  it('keys an explicit currentDate the same way the render path will', async () => {
    // A viewer with an explicit currentDate would otherwise miss the primed key
    // exactly as the tracked-changes viewer did.
    const dated = normalizeLayoutOptions(new Date(CURRENT_DATE_MS + 86_400_000), CURRENT_DATE_MS);
    const { source, services, store, builds } = spyStore('plain', 120);
    const layout = await layoutDocumentProgressively(
      source.bodyLayoutInput,
      services,
      dated,
    );
    store.prime(dated, layout);
    builds.length = 0;
    store.select(dated);
    expect(builds).toEqual([]);
  }, 300_000);
});
