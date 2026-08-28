import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { retainRenderWorkerDocumentLayout } from './render-worker-layout.js';
import {
  installStubCanvas,
  syntheticDocxModel,
  type SyntheticDocumentShape,
} from './testing/synthetic-document.js';
import { paginateBody } from './layout/body-paginator.js';
import { PaginationAbortError } from './layout/pagination-scheduler.js';
import { layoutFingerprint } from './layout/invariants.js';
import { layoutOptionsForRender, normalizeLayoutOptions } from './layout/options.js';
import { setDocumentLayoutValidation } from './layout/validation-policy.js';
import type { DocumentLayout } from './layout/types.js';
import {
  paginateRenderWorkerDocumentProgressively,
  type RenderWorkerLayoutPublication,
} from './render-worker-progressive.js';

// ─────────────────────────────────────────────────────────────────────────────
// Progressive layout INSIDE the render worker.
//
// Exercised through the same pure seam `render-worker.ts` calls, so none of
// this needs a Worker, an OffscreenCanvas or WASM — the trick
// `progressive-load.test.ts` already uses for the main-thread twin.
//
// The contract being pinned has three parts. Publications must grow, never
// shrink. Each publication must be PRIMED before it is announced, so a
// `renderPage` arriving on the very next task can paint the page it was just
// told about. And the layout left in the store at the end must be exactly the
// one a blocking parse would have produced — because `render-worker.ts` reads
// it straight back through `doc.layoutVariants.defaultLayout` to build the
// authoritative `parsedMeta`.
// ─────────────────────────────────────────────────────────────────────────────

const DEFAULT_CURRENT_DATE_MS = 1_700_000_000_000;
const DEFAULT_VIEW = layoutOptionsForRender({ defaultCurrentDateMs: DEFAULT_CURRENT_DATE_MS });
const MARKUP_VIEW = normalizeLayoutOptions(undefined, DEFAULT_CURRENT_DATE_MS, true);

function retain(shape: SyntheticDocumentShape, paragraphs: number) {
  const source = layoutSourceStore(syntheticDocxModel(shape, { paragraphs }));
  const services = createLayoutServices(source);
  const retained = retainRenderWorkerDocumentLayout(
    source,
    services,
    DEFAULT_CURRENT_DATE_MS,
  );
  return { source, services, retained };
}

/** Collect publications while recording what the store served at the instant
 *  each one was announced — the "prime before publish" invariant. */
function recordingPublisher(storePageCount: () => number) {
  const publications: RenderWorkerLayoutPublication[] = [];
  const servedAtPublish: number[] = [];
  const progress: number[] = [];
  return {
    publications,
    servedAtPublish,
    progress,
    publisher: {
      publish: (publication: RenderWorkerLayoutPublication) => {
        publications.push(publication);
        servedAtPublish.push(storePageCount());
      },
      progress: (committedPages: number) => { progress.push(committedPages); },
    },
  };
}

beforeAll(() => {
  installStubCanvas();
});

afterAll(() => {
  setDocumentLayoutValidation(true);
});

describe('render worker progressive layout', () => {
  it('primes each prefix before announcing it, and ends on the blocking layout', async () => {
    const { source, retained } = retain('plain', 300);
    const store = retained.layoutVariants;
    const recorder = recordingPublisher(() => store.defaultLayout.pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained,
      source,
      recorder.publisher,
      DEFAULT_VIEW,
    );

    expect(recorder.publications.length).toBeGreaterThan(0);
    // Prime-before-publish: the store already served exactly the announced
    // pages when the announcement went out. Anything less would let the host
    // request a page the worker cannot yet lay its hands on.
    recorder.publications.forEach((publication, index) => {
      expect(recorder.servedAtPublish[index]).toBe(publication.pageCount);
      expect(publication.pageSizes).toHaveLength(publication.pageCount);
      // Later convergence can still replace a checkpoint, so publications stay
      // provisional even though the canonical source is never truncated.
      expect(publication.exact).toBe(false);
    });
    // Monotonic: a shrinking page count would jump the viewport under a reader.
    const counts = recorder.publications.map((publication) => publication.pageCount);
    expect([...counts].sort((a, b) => a - b)).toEqual(counts);

    // The store is left holding the AUTHORITATIVE layout under the default key,
    // which is what makes `render-worker.ts`'s unchanged
    // `doc.layoutVariants.defaultLayout` a cache hit rather than a second pass.
    const fresh = layoutSourceStore(syntheticDocxModel('plain', { paragraphs: 300 }));
    const blocking = paginateBody(
      fresh.bodyLayoutInput,
      createLayoutServices(fresh),
      layoutOptionsForRender({ defaultCurrentDateMs: DEFAULT_CURRENT_DATE_MS }),
    );
    expect(layoutFingerprint(store.defaultLayout as DocumentLayout))
      .toBe(layoutFingerprint(blocking));
    expect(store.defaultLayout.pages.length)
      .toBeGreaterThan(recorder.publications.at(-1)!.pageCount);
  }, 300_000);

  it('cannot overwrite a newer same-key authority after losing publication ownership', async () => {
    const { source, retained } = retain('plain', 300);
    const store = retained.layoutVariants;
    // Stand in for a synchronous rebuild triggered by a view switch between
    // progressive slices. Its distinct pagination makes any stale overwrite
    // observable even though it occupies the same options key.
    const replacementSource = layoutSourceStore(syntheticDocxModel('plain', { paragraphs: 60 }));
    const replacement = paginateBody(
      replacementSource.bodyLayoutInput,
      createLayoutServices(replacementSource),
      DEFAULT_VIEW,
    );
    let publications = 0;

    await paginateRenderWorkerDocumentProgressively(
      retained,
      source,
      {
        publish: () => {
          publications += 1;
          if (publications !== 1) return;
          const progressivePrefix = store.layoutFor(DEFAULT_VIEW);
          expect(store.replaceIfCurrent(DEFAULT_VIEW, progressivePrefix, replacement))
            .not.toBeNull();
        },
        progress: () => {},
      },
      DEFAULT_VIEW,
    );

    expect(publications).toBe(1);
    expect(layoutFingerprint(store.layoutFor(DEFAULT_VIEW) as DocumentLayout))
      .toBe(layoutFingerprint(replacement));
  }, 300_000);

  it('reports progress so a silent worker is distinguishable from a busy one', async () => {
    // The host gives up its request timeout for the duration of a progressive
    // parse, so these are its only liveness evidence between publications.
    const { source, retained } = retain('plain', 300);
    const recorder = recordingPublisher(() => retained.layoutVariants.defaultLayout.pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained, source, recorder.publisher, DEFAULT_VIEW);

    expect(recorder.progress.length).toBeGreaterThan(0);
  }, 300_000);

  it('resolves bookmark anchors within the published prefix only', async () => {
    const { source, retained } = retain('plain', 300);
    const recorder = recordingPublisher(() => retained.layoutVariants.defaultLayout.pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained, source, recorder.publisher, DEFAULT_VIEW);

    for (const publication of recorder.publications) {
      // A prefix map may be empty, but it may never name a page that prefix
      // does not have — the host resolves internal hyperlinks against this.
      for (const [, pageIndex] of publication.bookmarkPages) {
        expect(pageIndex).toBeLessThan(publication.pageCount);
      }
    }
  }, 300_000);

  it('abandons the drain when the parse is superseded', async () => {
    // A re-parse aborts the previous document's drain. It must stop rather than
    // keep paginating for a document the worker has already dropped.
    const { source, retained } = retain('plain', 300);
    const abort = new AbortController();
    const store = retained.layoutVariants;
    let published = 0;

    const drain = paginateRenderWorkerDocumentProgressively(
      retained,
      source,
      {
        publish: () => { published += 1; abort.abort(); },
        progress: () => {},
      },
      DEFAULT_VIEW,
      abort.signal,
    );

    await expect(drain).rejects.toBeInstanceOf(PaginationAbortError);
    expect(published).toBe(1);
    // The prefix it managed to prime is still there; nothing half-written.
    expect(store.defaultLayout.pages.length).toBeGreaterThan(0);
  }, 300_000);

  it('previews and primes the variant the load selected, not the default one', async () => {
    // The markup view genuinely paginates differently — deletions stay visible,
    // so line breaking and page breaks move. Priming under the default key
    // would leave the progressive pass unread AND make the worker build a
    // second full layout for the view it actually paints.
    const source = layoutSourceStore(syntheticDocxModel('tracked-fields', { paragraphs: 200 }));
    const services = createLayoutServices(source);
    const retained = retainRenderWorkerDocumentLayout(source, services, DEFAULT_CURRENT_DATE_MS);
    const store = retained.layoutVariants;
    const recorder = recordingPublisher(() => store.layoutFor(MARKUP_VIEW).pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained, source, recorder.publisher, MARKUP_VIEW);

    expect(recorder.publications.length).toBeGreaterThan(0);
    // Every publication was primed under the markup key, so the store served
    // exactly the announced pages for THAT view.
    recorder.publications.forEach((publication, index) => {
      expect(recorder.servedAtPublish[index]).toBe(publication.pageCount);
    });

    // The markup layout is now cached, so the worker's metadata route reads it
    // back rather than paginating again...
    expect(store.hasLayoutFor(MARKUP_VIEW)).toBe(true);
    const markupPages = store.layoutFor(MARKUP_VIEW).pages.length;
    // ...and it is a genuinely different pagination from the final view, which
    // is what makes reporting the default variant a real bug rather than a
    // stylistic one.
    expect(store.layoutFor(DEFAULT_VIEW).pages.length).not.toBe(markupPages);

    const fresh = layoutSourceStore(syntheticDocxModel('tracked-fields', { paragraphs: 200 }));
    const blocking = paginateBody(
      fresh.bodyLayoutInput,
      createLayoutServices(fresh),
      MARKUP_VIEW,
    );
    expect(layoutFingerprint(store.layoutFor(MARKUP_VIEW) as DocumentLayout))
      .toBe(layoutFingerprint(blocking));
  }, 300_000);

  it('honours an explicit currentDate as its own variant', async () => {
    // DATE/TIME field text changes measured widths, so the date is an
    // acquisition input with its own pagination — not a paint-time detail.
    const { source, retained } = retain('fields', 200);
    const dated = normalizeLayoutOptions(DEFAULT_CURRENT_DATE_MS + 86_400_000 * 400, DEFAULT_CURRENT_DATE_MS);
    const store = retained.layoutVariants;
    const recorder = recordingPublisher(() => store.layoutFor(dated).pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained, source, recorder.publisher, dated);

    expect(store.hasLayoutFor(dated)).toBe(true);
    const fresh = layoutSourceStore(syntheticDocxModel('fields', { paragraphs: 200 }));
    const blocking = paginateBody(fresh.bodyLayoutInput, createLayoutServices(fresh), dated);
    expect(layoutFingerprint(store.layoutFor(dated) as DocumentLayout))
      .toBe(layoutFingerprint(blocking));
  }, 300_000);

  it('still deposits the authoritative layout when no preview is publishable', async () => {
    // A document short enough that previewing is pointless publishes nothing —
    // but the store must still end up primed, or the worker's metadata route
    // would pay for a second full layout.
    const { source, retained } = retain('plain', 6);
    const recorder = recordingPublisher(() => retained.layoutVariants.defaultLayout.pages.length);

    await paginateRenderWorkerDocumentProgressively(
      retained, source, recorder.publisher, DEFAULT_VIEW);

    expect(recorder.publications).toHaveLength(0);
    expect(retained.layoutVariants.defaultLayout.pages.length).toBeGreaterThan(0);
  }, 300_000);
});
