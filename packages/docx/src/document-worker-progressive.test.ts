import { describe, it, expect, vi } from 'vitest';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { normalizeLayoutOptions } from './layout/options.js';
import type {
  DocumentLayoutMeta,
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

/** The layout-derived half a `selectLayoutView` reply carries. */
function layoutMeta(pageCount: number): DocumentLayoutMeta {
  return {
    pageCount,
    pageSizes: Array.from({ length: pageCount }, () => ({ ...PAGE })),
    bookmarkPages: [['intro', 0], ['outro', pageCount - 1]],
    commentAnchorRanges: [],
    revisionAnchorRanges: [],
  };
}

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
  /** Pages the worker reports for a runtime `selectLayoutView`. */
  switchedPageCount?: number;
  onPartial?: (p: { pageCount: number; exact: boolean }) => void;
  onComplete?: (error?: unknown) => void;
  onProgress?: (p: { committedPages: number }) => void;
} = {}) {
  let settle!: (res: RenderWorkerResponse) => void;
  let fail!: (error: unknown) => void;
  const reply = new Promise<RenderWorkerResponse>((res, rej) => { settle = res; fail = rej; });
  const requests: RenderWorkerRequest[] = [];
  // The parse takes 11, matching what every settle() below correlates on.
  let nextId = 11;
  let terminated = false;

  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(document, {
    _mode: 'worker',
    _document: null,
    _source: null,
    _meta: null,
    _layoutComplete: true,
    // Field initializers the real constructor runs; destroy() reads them.
    _rawParts: new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 }),
    _embeddedFontFaces: [],
    _googleFontFaces: [],
    _localMetricFontFaces: [],
    _bridge: {
      request: (factory: (id: number) => RenderWorkerRequest) => {
        const request = factory(nextId++);
        requests.push(request);
        // A runtime view switch is its own round-trip; the parse's own reply
        // stays under the test's control.
        if (request.type === 'selectLayoutView') {
          return Promise.resolve({
            type: 'layoutViewSelected',
            id: request.id,
            meta: layoutMeta(opts.switchedPageCount ?? 13),
          } satisfies RenderWorkerResponse);
        }
        return reply;
      },
      terminate: () => { terminated = true; },
    },
  });
  attachDocumentLayoutRuntime(document, 0);
  if (opts.view) {
    // load() records the active variant before parsing; the parse request is
    // derived from that same record.
    // Straight to the runtime record, exactly as load() does before parsing —
    // the public setter is a worker round-trip and there is no worker yet.
    documentLayoutRuntimeOf(document).activeLayoutOptions = normalizeLayoutOptions(
      opts.view.currentDateMs,
      0,
      opts.view.showTrackedChanges === true,
    );
  }

  const progressive = {
    onPartial: opts.onPartial,
    onComplete: opts.onComplete,
    onProgress: opts.onProgress,
    abort: new AbortController(),
    firstPublication: (() => {
      let resolve!: () => void;
      let reject!: (e: unknown) => void;
      const promise = new Promise<void>((res, rej) => { resolve = res; reject = rej; });
      return { promise, resolve, reject };
    })(),
    published: false,
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
    const partials: { pageCount: number; exact: boolean }[] = [];
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
    const partials: { pageCount: number; exact: boolean }[] = [];
    let completed = 0;
    const harness = progressiveDocument({
      onPartial: (p) => partials.push(p),
      onComplete: () => { completed += 1; },
    });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(9) });

    expect(harness.document.pageCount).toBe(9);
    expect(partials).toEqual([{ pageCount: 9, exact: false }]);
    // Review data established by the first publication survives later ones,
    // which deliberately do not re-send it.
    expect(harness.document.comments).toHaveLength(1);
    // Anchor projections are whole-document joins; the worker omits them from a
    // prefix rather than ship truncated ones.
    expect(harness.document.commentAnchorRanges()).toEqual([]);
    // Identity-stable, so a per-frame consumer caching on identity does not
    // rebuild every draw.
    expect(harness.document.commentAnchorRanges())
      .toBe(harness.document.commentAnchorRanges());

    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(40) });
    await harness.document.whenLayoutComplete();

    expect(harness.document.pageCount).toBe(40);
    expect(harness.document.layoutComplete).toBe(true);
    expect(harness.document.getBookmarkPage('outro')).toBe(39);
    expect(completed).toBe(1);
  });

  it('keeps a switched view when the superseded parse keeps publishing', async () => {
    // Progressive layout resolves load() early, so a viewer can switch the
    // §17.13.5 view while the worker is still paginating the LOAD-TIME variant.
    // Its later prefixes and its authoritative metadata describe that variant;
    // installing either would put its page count and page geometry behind the
    // markup layout now being painted.
    const harness = progressiveDocument({ switchedPageCount: 13 });
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    expect(harness.document.pageCount).toBe(2);

    await harness.document.setLayoutView({ showTrackedChanges: true });
    expect(harness.document.pageCount).toBe(13);
    expect(harness.document.getBookmarkPage('outro')).toBe(12);

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(9) });
    expect(harness.document.pageCount).toBe(13);

    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(40) });
    await harness.document.whenLayoutComplete();

    expect(harness.document.pageCount).toBe(13);
    expect(harness.document.getBookmarkPage('outro')).toBe(12);
    // The progressive window still closes, and the model-derived records the
    // authoritative metadata carries ARE variant-independent, so they land.
    expect(harness.document.layoutComplete).toBe(true);
    expect(harness.document.comments).toHaveLength(1);
  });

  it('installs the load-time variant again when the view is switched back', async () => {
    // The guard is variant equality, not "a switch ever happened": returning to
    // the parse's own variant must let its metadata through again.
    const harness = progressiveDocument({ switchedPageCount: 13 });
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    await harness.document.setLayoutView({ showTrackedChanges: true });
    expect(harness.document.pageCount).toBe(13);
    await harness.document.setLayoutView({});
    // Switching back to the installed-at-load variant needs no round-trip, but
    // it does hand the parse its authority back.
    harness.settle({ type: 'parsedMeta', id: 11, meta: fullMeta(40) });
    await harness.document.whenLayoutComplete();

    expect(harness.document.pageCount).toBe(40);
    expect(harness.document.getBookmarkPage('outro')).toBe(39);
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
    const progress: { committedPages: number }[] = [];
    const harness = progressiveDocument({ onProgress: (p) => progress.push(p) });
    harness.push({ type: 'layoutProgress', forId: 11, committedPages: 17 });
    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;

    expect(progress).toEqual([{ committedPages: 17 }]);
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

  it('reports a failure arriving after load() resolved through whenLayoutComplete', async () => {
    const errors: unknown[] = [];
    const harness = progressiveDocument({ onComplete: (error) => errors.push(error) });

    harness.push({ type: 'layoutPartial', forId: 11, partial: partial(2, { review: REVIEW }) });
    await harness.parsed;
    harness.fail(new Error('background layout failed'));

    await expect(harness.document.whenLayoutComplete()).rejects.toThrow('background layout failed');
    expect(errors).toHaveLength(1);
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
    // tell, and whenLayoutComplete() must not reject for it.
    await expect(harness.document.whenLayoutComplete()).resolves.toBeUndefined();
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

      // Going quiet is what is not allowed. The message must name the deadline
      // it actually enforced — a diagnostic that reads `undefinedms` tells an
      // integrator nothing about which timeout to raise.
      vi.advanceTimersByTime(1_001);
      await expect(harness.parsed).rejects.toThrow(
        'worker layout produced no progress for 1000ms',
      );
      expect(harness.terminated()).toBe(true);
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
