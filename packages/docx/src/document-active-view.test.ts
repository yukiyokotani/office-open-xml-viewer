import { beforeAll, describe, expect, it } from 'vitest';
import { DocxDocument } from './document.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { layoutDocumentInput } from './layout/document.js';
import { normalizeLayoutOptions, type LayoutOptions } from './layout/options.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
} from './layout/runtime-state.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';

// ─────────────────────────────────────────────────────────────────────────────
// `load({ currentDate, showTrackedChanges })` primes and records the variant
// the caller will render so the first render does not synchronously
// repaginate. The bug this pins: a direct `DocxDocument` API call that OMITS
// the per-call option used to select the DEFAULT variant anyway — silently
// building a second full layout and letting paint disagree with the geometry
// accessors. Omitted options must now select the recorded active variant; an
// explicitly passed value (including `showTrackedChanges: false`) still wins.
//
// Exercised through `collectPageRuns`, which shares `_withActiveView` with
// `renderPage`, `renderPageToBitmap` and `getElementContextAt`. The builder
// spy is the proof: a wrong-variant selection is invisible in the output (it
// is correct, just enormously expensive and thrown away), so counting builds
// is the only way to state the guarantee.
// ─────────────────────────────────────────────────────────────────────────────

const CURRENT_DATE_MS = 1_700_000_000_000;
const DATED = new Date(CURRENT_DATE_MS + 86_400_000);

function documentWithActiveView(view: {
  currentDate?: Date | number;
  showTrackedChanges?: boolean;
}) {
  const source = layoutSourceStore(syntheticDocxModel('tracked', { paragraphs: 20 }));
  const services = createLayoutServices(source);
  const builds: LayoutOptions[] = [];
  attachDocumentLayoutVariants({
    source,
    services,
    defaultCurrentDateMs: CURRENT_DATE_MS,
    buildLayout: (options) => {
      builds.push(options);
      return layoutDocumentInput(source.bodyLayoutInput, services, options);
    },
  });
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, { _mode: 'main', _document: null, _source: null, _meta: null });
  attachDocumentLayoutRuntime(doc, CURRENT_DATE_MS);
  const runtime = documentLayoutRuntimeOf(doc);
  runtime.services = services;
  // What load() records for the variant it primes.
  runtime.activeLayoutOptions = normalizeLayoutOptions(
    view.currentDate,
    CURRENT_DATE_MS,
    view.showTrackedChanges === true,
  );
  return { doc, builds };
}

beforeAll(() => {
  installStubCanvas();
});

describe('omitted per-call options select the active load-time variant', () => {
  it('an optionless collectPageRuns never builds the default variant', async () => {
    const { doc, builds } = documentWithActiveView({ currentDate: DATED });
    await doc.collectPageRuns(0);
    // Exactly one build — the dated variant load() recorded — and no second,
    // default-keyed pagination behind the caller's back.
    expect(builds).toHaveLength(1);
    expect(builds[0]!.currentDateMs).toBe(DATED.getTime());
  }, 300_000);

  it('an optionless call follows the loaded markup view', async () => {
    const { doc, builds } = documentWithActiveView({ showTrackedChanges: true });
    await doc.collectPageRuns(0);
    expect(builds).toHaveLength(1);
    expect(builds[0]!.showTrackedChanges).toBe(true);
  }, 300_000);

  it('an explicit showTrackedChanges: false still selects the final view', async () => {
    const { doc, builds } = documentWithActiveView({ showTrackedChanges: true });
    await doc.collectPageRuns(0, { showTrackedChanges: false });
    expect(builds).toHaveLength(1);
    expect(builds[0]!.showTrackedChanges).toBeUndefined();
  }, 300_000);
});
