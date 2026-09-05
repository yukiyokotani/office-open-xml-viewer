import { describe, it, expect, vi } from 'vitest';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { ProgressiveLayoutObserverNotifier } from '@silurus/ooxml-core/internal/progressive-layout-observers';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { subscribeDocxLayout } from './document-layout-events.js';
import {
  selectDocxLayoutView,
  subscribeDocxLayoutView,
  type DocxLayoutViewPublication,
} from './document-layout-view.js';
import type {
  DocumentLayoutPartial,
  DocumentMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';

// ─────────────────────────────────────────────────────────────────────────────
// The HOST half of worker-mode progressive layout.
//
// In worker mode the model never crosses the wire, so everything the document
// can answer during the provisional window — page count, page sizes, bookmark
// anchors, comments — comes from metadata the worker pushes. These pin that
// state machine: which pushes are honoured, what `load()` waits for, and how a
// failure that arrives AFTER load() resolved is reported.
//
// Built off-prototype with an injected bridge (the established pattern from
// `render-worker-layout-parity.test.ts`), because the real constructor opens a
// Worker.
// ─────────────────────────────────────────────────────────────────────────────

const PAGE = { widthPt: 595, heightPt: 842 };

function partial(pageCount: number, over: Partial<DocumentLayoutPartial> = {}): DocumentLayoutPartial {
  return {
    pageCount,
    pageSizes: Array.from({ length: pageCount }, () => ({ ...PAGE })),
    bookmarkPages: [['intro', 0]],
    commentAnchorRanges: [],
    revisionAnchorRanges: [],
    exact: false,
    ...over,
  };
}

const REVIEW = {
  revisions: [],
  comments: [{ id: '1', author: 'A', initials: 'A', date: '', text: 'hello' }],
  footnotes: [],
  endnotes: [],
} as unknown as NonNullable<DocumentLayoutPartial['review']>;

function fullMeta(pageCount: number): DocumentMeta {
  return {
    pageCount,
    revisions: [],
    comments: REVIEW.comments,
    footnotes: [],
    endnotes: [],
    pageSizes: Array.from({ length: pageCount }, () => ({ ...PAGE })),
    bookmarkPages: [['intro', 0], ['outro', pageCount - 1]],
    commentAnchorRanges: [],
    revisionAnchorRanges: [],
  } as unknown as DocumentMeta;
}

/**
 * A `DocxDocument` in worker mode whose `parse` reply is under test control.
 * `push` delivers an uncorrelated worker message through the same
 * `onUnsolicited` route the real bridge uses.
 */
function progressiveDocument(opts: {
  timeoutMs?: number;
  view?: { currentDateMs?: number; showTrackedChanges?: boolean };
  onPartial?: (p: { availableUnits: number; exact: boolean }) => void;
  onComplete?: (error?: unknown) => void;
  onProgress?: (p: { committedUnits: number }) => void;
} = {}) {
  let settle!: (res: RenderWorkerResponse) => void;
  let fail!: (error: unknown) => void;
  const reply = new Promise<RenderWorkerResponse>((res, rej) => { settle = res; fail = rej; });
  const requests: RenderWorkerRequest[] = [];
  let terminated = false;

  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(document, {
    _mode: 'worker',
    _document: null,
    _source: null,
    _meta: null,
    _layoutLifecycle: new ProgressiveLayoutLifecycle(),
    _layoutObservers: new ProgressiveLayoutObserverNotifier(),
    _layoutViewGeneration: 0,
    // Field initializers the real constructor runs; destroy() reads them.
    _rawParts: new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 }),
    _embeddedFontFaces: [],
    _googleFontFaces: [],
    _localMetricFontFaces: [],
    _bridge: {
      request: (factory: (id: number) => RenderWorkerRequest) => {
        requests.push(factory(11));
        return reply;
      },
      terminate: () => { terminated = true; },
    },
  });
  attachDocumentLayoutRuntime(document, 0);
  documentLayoutRuntimeOf(document).activeLayoutOptions = {
    currentDateMs: 0,
    showTrackedChanges: false,
  };
  if (opts.view) {
    // load() records the active variant before parsing; the parse request is
    // derived from that same record.
    (document as unknown as { setLayoutView(v: unknown): void }).setLayoutView({
      currentDate: opts.view.currentDateMs,
      showTrackedChanges: opts.view.showTrackedChanges,
    });
  }

  const progressive = {
    onPartial: opts.onPartial,
    onComplete: opts.onComplete,
    onProgress: opts.onProgress,
    layoutOptions: documentLayoutRuntimeOf(document).activeLayoutOptions,
    abort: new AbortController(),
    firstPublication: (() => {
      let resolve!: () => void;
      let reject!: (e: unknown) => void;
      const promise = new Promise<void>((res, rej) => { resolve = res; reject = rej; });
      return { promise, resolve, reject };
    })(),
    published: false,
    settled: false,
  };

  const parsed = (document as unknown as {
    _parse(
      buffer: ArrayBuffer,
      policy: unknown,
      google: boolean,
      timeoutMs: number | undefined,
      onUsage: unknown,
      renderers: unknown,
      progressive: unknown,
    ): Promise<void>;
  })._parse(new ArrayBuffer(1), undefined, false, opts.timeoutMs, undefined, undefined, progressive);

  const push = (res: RenderWorkerResponse): void => {
    (document as unknown as {
      _onWorkerLayoutPush(res: RenderWorkerResponse): void;
    })._onWorkerLayoutPush(res);
  };

  return {
    document,
    parsed,
    push,
    requests,
    settle,
    fail,
    terminated: () => terminated,
  };
}

describe('worker-mode progressive load', () => {
  it('asks the worker for progressive layout and resolves on the first publication', async () => {
    const partials: { availableUnits: number; exact: boolean }[] = [];
    const harness = progressiveDocument({ onPartial: (p) => partials.push(p) });

    const parseRequest = harness.requests[0];
    expect(parseRequest?.type).toBe('parse');
    expect(parseRequest && 'progressiveLayout' in parseRequest && parseRequest.progressiveLayout)
      .toBe(true);

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    // load() has returned on two real pages while the worker keeps paginating.
    expect(harness.document.pageCount).toBe(2);
    expect(harness.document.pageSize(1)).toEqual(PAGE);
    expect(harness.document.layoutComplete).toBe(false);
    // The model-derived review data rode along on the first publication, so the
    // document is not falsely empty during the provisional window.
    expect(harness.document.comments).toHaveLength(1);
    expect(harness.document.getBookmarkPage('intro')).toBe(0);
    // The first publication IS the loaded document, not an extension of one.
    expect(partials).toHaveLength(0);
  });

  it('grows page geometry on later publications and settles on the authoritative meta', async () => {
    const partials: { availableUnits: number; exact: boolean }[] = [];
    let completed = 0;
    const harness = progressiveDocument({
      onPartial: (p) => partials.push(p),
      onComplete: () => { completed += 1; },
    });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(9) });

    expect(harness.document.pageCount).toBe(9);
    expect(partials).toEqual([{ availableUnits: 9, exact: false }]);
    // Review data established by the first publication survives later ones,
    // which deliberately do not re-send it.
    expect(harness.document.comments).toHaveLength(1);
    // Prefix projections are authoritative for the pages already published.
    expect(harness.document.commentAnchorRanges()).toEqual([]);
    // Identity-stable, so a per-frame consumer caching on identity does not
    // rebuild every draw.
    expect(harness.document.commentAnchorRanges())
      .toBe(harness.document.commentAnchorRanges());

    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(40) });
    await harness.document.waitUntilLayoutComplete();

    expect(harness.document.pageCount).toBe(40);
    expect(harness.document.layoutComplete).toBe(true);
    expect(harness.document.getBookmarkPage('outro')).toBe(39);
    expect(completed).toBe(1);
  });

  it('reports completion exactly once when the worker publishes nothing (fast document)', async () => {
    // A document short enough to finish before the first checkpoint push never
    // sends `layoutPartial`: load() resolves on the authoritative meta itself.
    // The terminal callback contract must not depend on that — the consumer
    // registered for completion either way.
    const completions: unknown[] = [];
    const harness = progressiveDocument({
      onComplete: (error) => completions.push(error),
    });

    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(2) });
    await harness.parsed;
    await harness.document.waitUntilLayoutComplete();

    expect(harness.document.pageCount).toBe(2);
    expect(harness.document.layoutComplete).toBe(true);
    expect(completions).toEqual([undefined]);
  });

  it('reports completion exactly once after partials (no double-fire)', async () => {
    const completions: unknown[] = [];
    const harness = progressiveDocument({
      onComplete: (error) => completions.push(error),
    });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(9) });
    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(40) });
    await harness.document.waitUntilLayoutComplete();

    expect(completions).toEqual([undefined]);
  });

  it('sends the default view as no view fields at all', async () => {
    // Keeps the wire shape identical to what pre-variant builds sent, so a
    // default load cannot accidentally select a different key.
    const harness = progressiveDocument();
    const parse = harness.requests[0];

    expect(parse && 'currentDateMs' in parse).toBe(false);
    expect(parse && 'showTrackedChanges' in parse).toBe(false);
  });

  it('carries the selected variant to the worker so metadata describes the painted view', async () => {
    // Before this, a worker-mode markup load reported the FINAL view's page
    // count while painting the markup one — the two genuinely differ.
    const harness = progressiveDocument({ view: { showTrackedChanges: true } });
    const parse = harness.requests[0];

    expect(parse && 'showTrackedChanges' in parse && parse.showTrackedChanges).toBe(true);
  });

  it('carries an explicit currentDate as a variant axis', async () => {
    const harness = progressiveDocument({ view: { currentDateMs: 5_000 } });
    const parse = harness.requests[0];

    expect(parse && 'currentDateMs' in parse && parse.currentDateMs).toBe(5_000);
  });

  it('ignores a push naming a parse this document has moved past', async () => {
    const harness = progressiveDocument();
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    harness.push({ type: 'layoutPartial', forId: 99, partial: partial(500) });

    expect(harness.document.pageCount).toBe(2);
  });

  it('forwards throttled worker progress', async () => {
    const progress: { committedUnits: number }[] = [];
    const harness = progressiveDocument({ onProgress: (p) => progress.push(p) });
    harness.push({ type: 'layoutProgress', forId: 11, committedPages: 17 });
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    expect(progress).toEqual([{ committedUnits: 17 }]);
  });

  it('keeps observer exceptions out of the authoritative layout result', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const harness = progressiveDocument({
      onProgress: () => { throw new Error('progress observer failed'); },
      onPartial: () => { throw new Error('partial observer failed'); },
      onComplete: () => { throw new Error('complete observer failed'); },
    });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.push({ type: 'layoutProgress', forId: 11, committedPages: 3 });
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(4) });
    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(5) });

    await expect(harness.document.waitUntilLayoutComplete()).resolves.toBeUndefined();
    expect(harness.document.layoutComplete).toBe(true);
    expect(harness.document.pageCount).toBe(5);
    expect(consoleError).toHaveBeenCalledTimes(3);
    consoleError.mockRestore();
  });

  it('rejects load() when the worker fails before publishing anything', async () => {
    let completed = 0;
    const harness = progressiveDocument({ onComplete: () => { completed += 1; } });

    harness.fail(new Error('worker exploded'));

    await expect(harness.parsed).rejects.toThrow('worker exploded');
    // Nothing was shown early, so this is still load()'s own rejection — not a
    // background failure the caller has to go looking for.
    expect(completed).toBe(0);
  });

  it('reports a failure arriving after load() resolved through waitUntilLayoutComplete', async () => {
    const errors: unknown[] = [];
    const harness = progressiveDocument({ onComplete: (error) => errors.push(error) });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.fail(new Error('background layout failed'));

    await expect(harness.document.waitUntilLayoutComplete()).rejects.toThrow('background layout failed');
    expect(errors).toHaveLength(1);
    expect(harness.document.layoutComplete).toBe(false);
    // The provisional pages stay usable; only the completion is lost.
    expect(harness.document.pageCount).toBe(2);
  });

  it('settles quietly when the document is destroyed mid-layout', async () => {
    let completed = 0;
    const harness = progressiveDocument({ onComplete: () => { completed += 1; } });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    harness.document.destroy();
    harness.fail(new Error('Worker terminated'));

    // A deliberate teardown is not a layout failure: there is nobody left to
    // tell, and waitUntilLayoutComplete() must not reject for it.
    await expect(harness.document.waitUntilLayoutComplete()).resolves.toBeUndefined();
    expect(completed).toBe(0);
  });

  it('treats worker silence, not total elapsed time, as the failure condition', async () => {
    vi.useFakeTimers();
    try {
      const harness = progressiveDocument({ timeoutMs: 1_000 });

      // A background layout may legitimately outlive any fixed deadline, so
      // long as it keeps saying so.
      for (let elapsed = 0; elapsed < 5_000; elapsed += 900) {
        vi.advanceTimersByTime(900);
        harness.push({ type: 'layoutProgress', forId: 11, committedPages: elapsed });
      }
      expect(harness.terminated()).toBe(false);

      // Going quiet is what is not allowed.
      vi.advanceTimersByTime(1_001);
      await expect(harness.parsed).rejects.toThrow(
        'worker layout produced no progress for 1000ms',
      );
      expect(harness.terminated()).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });

  it('keeps a post-publication silence failure as the single terminal result', async () => {
    vi.useFakeTimers();
    try {
      const completions: unknown[] = [];
      const publications: unknown[] = [];
      const harness = progressiveDocument({
        timeoutMs: 1_000,
        onComplete: (error) => completions.push(error),
      });
      const unsubscribe = subscribeDocxLayout(
        harness.document,
        () => ({ pageCount: harness.document.pageCount, exact: false, complete: false }),
        (publication) => {
          if (publication.error) publications.push(publication.error);
        },
        () => {},
      );
      harness.push({
        type: 'layoutPartial',
        forId: 11,
        partial: partial(2, { review: REVIEW }),
      });
      await harness.parsed;

      vi.advanceTimersByTime(1_001);
      expect(harness.terminated()).toBe(true);
      const silenceError = completions[0];
      expect(silenceError).toBeInstanceOf(Error);
      harness.fail(new Error('Worker terminated'));
      await Promise.resolve();

      await expect(harness.document.waitUntilLayoutComplete()).rejects.toBe(silenceError);
      expect(completions).toEqual([silenceError]);
      expect(publications).toEqual([silenceError]);
      unsubscribe();
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('progressive pushes and request correlation', () => {
  /** In-memory worker whose replies the test drives directly. */
  class ScriptedWorker implements WorkerLike {
    listeners: ((e: MessageEvent) => void)[] = [];
    postMessage(): void {}
    addEventListener(type: 'message', listener: (e: MessageEvent) => void): void;
    addEventListener(type: 'messageerror', listener: (e: MessageEvent) => void): void;
    addEventListener(type: 'error', listener: (e: ErrorEvent) => void): void;
    addEventListener(type: string, listener: (e: never) => void): void {
      if (type === 'message') this.listeners.push(listener as (e: MessageEvent) => void);
    }
    removeEventListener(): void {}
    terminate(): void {}
    emit(data: unknown): void {
      for (const listener of this.listeners) listener({ data } as MessageEvent);
    }
  }

  it('routes a forId push to onUnsolicited without settling the pending parse', async () => {
    // The whole mechanism rests on this: `correlate` keys on `id`, so a push
    // keyed on `forId` must NOT resolve the in-flight parse. If it did, the
    // authoritative `parsedMeta` would arrive with nowhere to go and the
    // document would be frozen at its preview prefix forever.
    const worker = new ScriptedWorker();
    const unsolicited: unknown[] = [];
    const bridge = new WorkerBridge<RenderWorkerResponse>(worker, {
      correlate: (res) => ('id' in res ? res.id : undefined),
      onUnsolicited: (res) => { unsolicited.push(res); },
    });

    let settled = false;
    const parse = bridge.request((id) => ({ type: 'parse', id })).then((res) => {
      settled = true;
      return res;
    });

    worker.emit({ type: 'layoutPartial', forId: 1, partial: partial(2) });
    worker.emit({ type: 'layoutProgress', forId: 1, committedPages: 5 });
    await Promise.resolve();

    expect(unsolicited).toHaveLength(2);
    expect(settled).toBe(false);

    worker.emit({ type: 'parsedMeta', id: 1, meta: fullMeta(40) });
    await expect(parse).resolves.toMatchObject({ type: 'parsedMeta' });
    expect(unsolicited).toHaveLength(2);
  });
});

describe('worker layout-view metadata switch', () => {
  it('publishes only the winning concurrent view after matching geometry is installed', async () => {
    const resolvers: Array<(value: RenderWorkerResponse) => void> = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _document: null,
      _source: null,
      _meta: fullMeta(11),
      _layoutViewGeneration: 0,
      _bridge: {
        request: () => new Promise<RenderWorkerResponse>((resolve) => resolvers.push(resolve)),
      },
    });
    attachDocumentLayoutRuntime(document, 0);
    documentLayoutRuntimeOf(document).activeLayoutOptions = {
      currentDateMs: 0,
      showTrackedChanges: false,
    };
    const publications: DocxLayoutViewPublication[] = [];
    const unsubscribe = subscribeDocxLayoutView(
      document,
      (publication) => publications.push(publication),
      vi.fn(),
    );

    const staleRequester = {};
    const winningRequester = {};
    const stale = selectDocxLayoutView(document, {
      showTrackedChanges: true,
      currentDate: 10,
    }, staleRequester);
    const winning = selectDocxLayoutView(document, {
      showTrackedChanges: true,
      currentDate: 20,
    }, winningRequester);
    expect(resolvers).toHaveLength(2);

    resolvers[0]!({
      type: 'layoutViewSelected',
      id: 1,
      meta: fullMeta(13),
    } as RenderWorkerResponse);
    await expect(stale).resolves.toBe(false);
    expect(publications).toEqual([]);

    resolvers[1]!({
      type: 'layoutViewSelected',
      id: 2,
      meta: fullMeta(14),
    } as RenderWorkerResponse);
    await expect(winning).resolves.toBe(true);
    expect(document.pageCount).toBe(14);
    expect(publications).toEqual([{
      view: { showTrackedChanges: true, currentDate: 20 },
      generation: 1,
      requester: winningRequester,
    }]);
    unsubscribe();
  });

  it('keeps the old variant active until matching worker geometry is ready', async () => {
    let resolveMeta!: (value: RenderWorkerResponse) => void;
    const metaResponse = new Promise<RenderWorkerResponse>((resolve) => {
      resolveMeta = resolve;
    });
    const requests: RenderWorkerRequest[] = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _document: null,
      _source: null,
      _meta: fullMeta(11),
      _layoutViewGeneration: 0,
      _bridge: {
        request: (factory: (id: number) => RenderWorkerRequest) => {
          requests.push(factory(23));
          return metaResponse;
        },
      },
    });
    attachDocumentLayoutRuntime(document, 0);
    const runtime = documentLayoutRuntimeOf(document);
    runtime.activeLayoutOptions = { currentDateMs: 0, showTrackedChanges: false };

    const switching = Promise.resolve(document.setLayoutView({ showTrackedChanges: true }));

    // Until the worker has paginated the requested variant, both synchronous
    // geometry and option fill-in must remain on the installed final view.
    expect(document.pageCount).toBe(11);
    expect(requests).toEqual([{
      type: 'selectLayoutView',
      id: 23,
      currentDateMs: 0,
      showTrackedChanges: true,
    }]);

    resolveMeta({
      type: 'layoutViewSelected',
      id: 23,
      meta: {
        pageCount: 13,
        pageSizes: Array.from({ length: 13 }, () => ({ ...PAGE })),
        bookmarkPages: [['outro', 12]],
        commentAnchorRanges: [],
        revisionAnchorRanges: [],
      },
    } as unknown as RenderWorkerResponse);
    await switching;

    expect(document.pageCount).toBe(13);
    expect(document.pageSize(12)).toEqual(PAGE);
  });

  it('a request for the installed view cancels an older in-flight switch', async () => {
    let resolveMarkup!: (value: RenderWorkerResponse) => void;
    const markupResponse = new Promise<RenderWorkerResponse>((resolve) => {
      resolveMarkup = resolve;
    });
    const requests: RenderWorkerRequest[] = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _document: null,
      _source: null,
      _meta: fullMeta(11),
      _layoutViewGeneration: 0,
      _bridge: {
        request: (factory: (id: number) => RenderWorkerRequest) => {
          requests.push(factory(29));
          return markupResponse;
        },
      },
    });
    attachDocumentLayoutRuntime(document, 0);
    documentLayoutRuntimeOf(document).activeLayoutOptions = {
      currentDateMs: 0,
      showTrackedChanges: false,
    };

    const markup = document.setLayoutView({ showTrackedChanges: true });
    await document.setLayoutView({ showTrackedChanges: false });
    resolveMarkup({
      type: 'layoutViewSelected',
      id: 29,
      meta: {
        pageCount: 13,
        pageSizes: Array.from({ length: 13 }, () => ({ ...PAGE })),
        bookmarkPages: [],
        commentAnchorRanges: [],
        revisionAnchorRanges: [],
      },
    } as unknown as RenderWorkerResponse);
    await markup;

    expect(requests).toHaveLength(1);
    expect(document.pageCount).toBe(11);
  });

  it('does not let the original progressive parse overwrite a selected variant', async () => {
    const selectedPage = { widthPt: 700, heightPt: 900 };
    let settleParse!: (value: RenderWorkerResponse) => void;
    const parseResponse = new Promise<RenderWorkerResponse>((resolve) => {
      settleParse = resolve;
    });
    const requests: RenderWorkerRequest[] = [];
    const document = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(document, {
      _mode: 'worker',
      _document: null,
      _source: null,
      _meta: null,
      _layoutLifecycle: new ProgressiveLayoutLifecycle(),
      _layoutObservers: new ProgressiveLayoutObserverNotifier(),
      _layoutViewGeneration: 0,
      _parseRequestId: null,
      _bridge: {
        request: (factory: (id: number) => RenderWorkerRequest) => {
          const request = factory(requests.length + 31);
          requests.push(request);
          if (request.type === 'selectLayoutView') {
            return Promise.resolve({
              type: 'layoutViewSelected',
              id: request.id,
              meta: {
                pageCount: 13,
                pageSizes: Array.from({ length: 13 }, () => ({ ...selectedPage })),
                bookmarkPages: [['markup-end', 12]],
                commentAnchorRanges: [],
                revisionAnchorRanges: [],
              },
            } satisfies RenderWorkerResponse);
          }
          return parseResponse;
        },
        terminate: () => undefined,
      },
    });
    attachDocumentLayoutRuntime(document, 0);
    const runtime = documentLayoutRuntimeOf(document);
    runtime.activeLayoutOptions = { currentDateMs: 0, showTrackedChanges: false };
    const progressive = {
      onPartial: vi.fn(),
      layoutOptions: runtime.activeLayoutOptions,
      abort: new AbortController(),
      firstPublication: (() => {
        let resolve!: () => void;
        let reject!: (error: unknown) => void;
        const promise = new Promise<void>((res, rej) => { resolve = res; reject = rej; });
        return { promise, resolve, reject };
      })(),
      published: false,
    };
    const loading = (document as unknown as {
      _parse(
        buffer: ArrayBuffer,
        policy: unknown,
        google: boolean,
        timeoutMs: number | undefined,
        onUsage: unknown,
        renderers: unknown,
        progressive: unknown,
      ): Promise<void>;
    })._parse(
      new ArrayBuffer(1),
      undefined,
      false,
      undefined,
      undefined,
      undefined,
      progressive,
    );

    (document as unknown as {
      _onWorkerLayoutPush(response: RenderWorkerResponse): void;
    })._onWorkerLayoutPush({
      type: 'layoutPartial',
      forId: 31,
      partial: partial(2, { review: REVIEW }),
    });
    await loading;
    await document.setLayoutView({ showTrackedChanges: true });
    expect(document.pageCount).toBe(13);
    expect(document.pageSize(12)).toEqual(selectedPage);
    expect(document.getBookmarkPage('markup-end')).toBe(12);

    // These messages still belong to the live parse request, but describe its
    // original final-view variant rather than the now-selected markup variant.
    (document as unknown as {
      _onWorkerLayoutPush(response: RenderWorkerResponse): void;
    })._onWorkerLayoutPush({
      type: 'layoutPartial',
      forId: 31,
      partial: partial(7),
    });
    expect(document.pageCount).toBe(13);
    expect(progressive.onPartial).not.toHaveBeenCalled();

    settleParse({ type: 'parsedMeta', id: 31, meta: fullMeta(40) });
    await document.waitUntilLayoutComplete();
    expect(document.layoutComplete).toBe(true);
    expect(document.pageCount).toBe(13);
    expect(document.pageSize(12)).toEqual(selectedPage);
    expect(document.getBookmarkPage('markup-end')).toBe(12);
    expect(document.getBookmarkPage('outro')).toBeUndefined();
  });
});
