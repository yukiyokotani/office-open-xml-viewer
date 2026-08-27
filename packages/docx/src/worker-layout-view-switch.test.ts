import { beforeAll, describe, expect, it } from 'vitest';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { DocxDocument } from './document.js';
import { buildBookmarkPageMap } from './bookmark-nav.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { normalizeLayoutOptions } from './layout/options.js';
import {
  attachDocumentLayoutRuntime,
  documentLayoutRuntimeOf,
} from './layout/runtime-state.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import {
  renderWorkerLayoutViewMeta,
  renderWorkerReviewAnchors,
} from './render-worker-layout-view.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
import type {
  DocumentLayoutMeta,
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol.js';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.13.5 view switching in `mode: 'worker'`, over the real protocol
// and a real pagination.
//
// The model never crosses the wire, so `pageCount`, `pageSize` and the anchor
// projections all come from worker metadata. Moving the active view without
// bringing the selected variant's metadata back therefore leaves the host
// painting the markup layout while measuring the final one — clamping, scroll
// extent and mount decisions all read the wrong pagination. A fake engine
// cannot show this: it is handed one page count and reports it for both views.
//
// The document below is the `tracked` synthetic fixture, whose two views
// genuinely paginate differently (pinned by `testing/tracked-fixture.test.ts`).
// ─────────────────────────────────────────────────────────────────────────────

const DEFAULT_DATE_MS = 1_700_000_000_000;
/** Smallest size at which the two views paginate to different page counts, so
 *  the suite pays for real pagination without paying for a long document. */
const PARAGRAPHS = 50;

beforeAll(() => {
  installStubCanvas();
});

/** The worker half: one retained document plus the review records the anchor
 *  projections need, exactly as `render-worker.ts` keeps them across requests.
 *
 *  Built once: the variant store caches each pagination, and every test here
 *  reads the same document a real worker would keep across requests. */
let shared: ReturnType<typeof buildTrackedDocument> | null = null;

function buildTrackedDocument() {
  const model = syntheticDocxModel('tracked', { paragraphs: PARAGRAPHS });
  const source = layoutSourceStore(model);
  const services = createLayoutServices(source);
  const retained = retainRenderWorkerDocumentLayout(source, services, DEFAULT_DATE_MS);
  const review = { comments: model.comments ?? [], revisions: model.revisions ?? [] };
  return { model, source, retained, review };
}

function retainedTrackedDocument(): ReturnType<typeof buildTrackedDocument> {
  shared ??= buildTrackedDocument();
  return shared;
}

function viewMeta(showTrackedChanges: boolean): DocumentLayoutMeta {
  const { retained, review } = retainedTrackedDocument();
  return renderWorkerLayoutViewMeta(retained, review, { showTrackedChanges });
}

describe('render worker — selectLayoutView metadata', () => {
  it('reports a different pagination for each tracked-changes view', () => {
    const { retained, review } = retainedTrackedDocument();

    const finalView = renderWorkerLayoutViewMeta(retained, review);
    const markupView = renderWorkerLayoutViewMeta(retained, review, { showTrackedChanges: true });

    // The markup view keeps deleted text, so it needs more pages. Equal counts
    // would mean this metadata could not tell a viewer which layout it is on.
    expect(markupView.pageCount).toBeGreaterThan(finalView.pageCount);
    expect(finalView.pageSizes).toHaveLength(finalView.pageCount);
    expect(markupView.pageSizes).toHaveLength(markupView.pageCount);
  }, 300_000);

  it('describes a variant exactly the way the parse route describes it', () => {
    // `parse` builds its metadata inline (its route is pinned by the DOCX
    // layout-boundary gate) and the switch builds it here. Both must be the
    // same projection of the same retained variant, or a switched view would
    // silently answer geometry questions differently from a load.
    const { retained, source, review } = retainedTrackedDocument();
    const layoutOptions = normalizeLayoutOptions(undefined, DEFAULT_DATE_MS, true);
    const layout = retained.layoutVariants.layoutFor(layoutOptions);
    const parseRoute: DocumentLayoutMeta = {
      pageCount: layout.pages.length,
      pageSizes: layout.pages.map((page) => ({
        widthPt: page.geometry.widthPt,
        heightPt: page.geometry.heightPt,
      })),
      bookmarkPages: [...buildBookmarkPageMap(layout)],
      ...renderWorkerReviewAnchors(layout, source, review),
    };

    expect(renderWorkerLayoutViewMeta(retained, review, { showTrackedChanges: true }))
      .toEqual(parseRoute);
  }, 300_000);
});

/**
 * A worker-mode `DocxDocument` whose bridge answers `selectLayoutView` from the
 * real worker seam over a real pagination. Built off-prototype with an injected
 * bridge (the established pattern here), because the real constructor opens a
 * Worker.
 */
function workerDocument(options: { defer?: boolean } = {}) {
  const { model, retained, review } = retainedTrackedDocument();
  const loadView = normalizeLayoutOptions(undefined, DEFAULT_DATE_MS, false);
  const loadMeta: DocumentMeta = {
    ...renderWorkerLayoutViewMeta(retained, review),
    revisions: model.revisions ?? [],
    comments: model.comments ?? [],
    footnotes: model.footnotes ?? [],
    endnotes: model.endnotes ?? [],
  };
  const requests: RenderWorkerRequest[] = [];
  const pending: Array<() => void> = [];
  let nextId = 1;

  const answer = (req: RenderWorkerRequest): RenderWorkerResponse => {
    if (req.type === 'selectLayoutView') {
      return {
        type: 'layoutViewSelected',
        id: req.id,
        meta: renderWorkerLayoutViewMeta(retained, review, {
          currentDateMs: req.currentDateMs,
          showTrackedChanges: req.showTrackedChanges,
        }),
      };
    }
    if (req.type === 'collectRuns') return { type: 'runsCollected', id: req.id, runs: [] };
    throw new Error(`unexpected worker request: ${req.type}`);
  };

  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(document, {
    _mode: 'worker',
    _document: null,
    _source: null,
    _meta: loadMeta,
    _layoutComplete: true,
    // Field initializers the real constructor runs; destroy() reads them.
    _rawParts: new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 }),
    _embeddedFontFaces: [],
    _googleFontFaces: [],
    _localMetricFontFaces: [],
    _bridge: {
      request: (factory: (id: number) => RenderWorkerRequest) => {
        const req = factory(nextId++);
        requests.push(req);
        if (!options.defer || req.type !== 'selectLayoutView') {
          return Promise.resolve(answer(req));
        }
        return new Promise<RenderWorkerResponse>((resolve) => {
          pending.push(() => { resolve(answer(req)); });
        });
      },
      terminate: () => {},
    },
  });
  attachDocumentLayoutRuntime(document, DEFAULT_DATE_MS);
  // load() records the variant it selected before any geometry is read.
  documentLayoutRuntimeOf(document).activeLayoutOptions = loadView;

  return { document, requests, pending, loadMeta };
}

describe('DocxDocument.setLayoutView — worker mode', () => {
  it('installs the selected variant’s geometry with the view itself', async () => {
    const { document, loadMeta } = workerDocument();
    const markup = viewMeta(true);
    expect(document.pageCount).toBe(loadMeta.pageCount);

    await document.setLayoutView({ showTrackedChanges: true });

    // Page count, per-page geometry and bookmark pages all follow the layout
    // the renderer is about to paint.
    expect(document.pageCount).toBe(markup.pageCount);
    expect(document.pageCount).not.toBe(loadMeta.pageCount);
    expect(document.pageSize(markup.pageCount - 1))
      .toEqual(markup.pageSizes[markup.pageCount - 1]);
    // Model-derived records are variant-independent, so the switch keeps the
    // ones the load established rather than re-shipping them.
    expect(document.comments).toEqual(loadMeta.comments);
    expect(document.footnotes).toEqual(loadMeta.footnotes);
  }, 300_000);

  it('answers for the previous variant until the switch lands', async () => {
    // The whole point of the round-trip: a reader between the request and the
    // reply must see one self-consistent view, not the new page count against
    // the old pages or the reverse.
    const { document, loadMeta, pending } = workerDocument({ defer: true });
    const markup = viewMeta(true);

    const switching = document.setLayoutView({ showTrackedChanges: true });
    await Promise.resolve();
    expect(document.pageCount).toBe(loadMeta.pageCount);
    expect(document.pageSize(0)).toEqual(loadMeta.pageSizes[0]);
    // And a render dispatched now still names the variant that geometry
    // describes.
    const { requests } = await collectRunsRequest(document);
    expect(requests.showTrackedChanges).toBeUndefined();

    expect(pending).toHaveLength(1);
    pending[0]!();
    await switching;

    expect(document.pageCount).toBe(markup.pageCount);
    const after = await collectRunsRequest(document);
    expect(after.requests.showTrackedChanges).toBe(true);
  }, 300_000);

  it('lets the newest selection win when two switches overlap', async () => {
    const { document, pending, requests, loadMeta } = workerDocument({ defer: true });
    const markup = viewMeta(true);

    const toMarkup = document.setLayoutView({ showTrackedChanges: true });
    const toDated = document.setLayoutView({ currentDate: DEFAULT_DATE_MS + 86_400_000 });
    expect(requests).toHaveLength(2);

    // The newest reply lands first, then the superseded one. The stale reply
    // must not overwrite the view that won.
    pending[1]!();
    pending[0]!();
    await Promise.all([toMarkup, toDated]);

    expect(document.pageCount).toBe(loadMeta.pageCount);
    expect(document.pageCount).not.toBe(markup.pageCount);
    const after = await collectRunsRequest(document);
    expect(after.requests.showTrackedChanges).toBeUndefined();
    expect(after.requests.currentDate).toBe(DEFAULT_DATE_MS + 86_400_000);
  }, 300_000);

  it('cancels an in-flight switch that is taken back before it lands', async () => {
    // Toggle on, toggle off. The second call selects the view that is already
    // installed, so there is nothing to fetch — but the first switch is still
    // in flight and must not install the markup view behind the reader.
    const { document, pending, loadMeta } = workerDocument({ defer: true });

    const toMarkup = document.setLayoutView({ showTrackedChanges: true });
    const backToFinal = document.setLayoutView({ showTrackedChanges: false });
    pending[0]!();
    await Promise.all([toMarkup, backToFinal]);

    expect(document.pageCount).toBe(loadMeta.pageCount);
    const after = await collectRunsRequest(document);
    expect(after.requests.showTrackedChanges).toBeUndefined();
  }, 300_000);

  it('joins a repeat request for the switch already in flight', async () => {
    const { document, pending, requests } = workerDocument({ defer: true });
    const markup = viewMeta(true);

    const first = document.setLayoutView({ showTrackedChanges: true });
    const second = document.setLayoutView({ showTrackedChanges: true });
    // One pagination, not two: the second call joins the first.
    expect(requests.filter((req) => req.type === 'selectLayoutView')).toHaveLength(1);

    pending[0]!();
    await Promise.all([first, second]);
    expect(document.pageCount).toBe(markup.pageCount);
  }, 300_000);
});

/** Dispatch a run collection and hand back the wire options it carried — the
 *  observable proof that renders are filled from the installed view. */
async function collectRunsRequest(
  document: DocxDocument,
): Promise<{ requests: { showTrackedChanges?: boolean; currentDate?: Date | number } }> {
  const seen: RenderWorkerRequest[] = [];
  const bridge = (document as unknown as {
    _bridge: { request(factory: (id: number) => RenderWorkerRequest): Promise<RenderWorkerResponse> };
  })._bridge;
  const original = bridge.request.bind(bridge);
  bridge.request = (factory) => {
    const captured = original((id) => {
      const req = factory(id);
      seen.push(req);
      return req;
    });
    return captured;
  };
  try {
    await document.collectPageRuns(0);
  } finally {
    bridge.request = original;
  }
  const request = seen.find((req) => req.type === 'collectRuns');
  if (!request || request.type !== 'collectRuns') throw new Error('no collectRuns request');
  return { requests: request.opts };
}
