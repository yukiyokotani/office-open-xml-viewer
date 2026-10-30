/**
 * Main-thread entry of the `fontResolver` worker bridge for one document
 * (`resolver-bridge.ts`). It recognises bridge messages synchronously and
 * loads the responder only on the document's first resolver request, so a
 * document without a `fontResolver` source, or one that paints on main, loads
 * no bridge code beyond this class. Messages are handled in arrival order
 * (a cancellation never overtakes its request). Each request's deadline
 * starts when it arrives here, so the time the responder takes to load counts
 * against it, as main mode's ceiling would; a request that expires before the
 * responder takes it is failed, recorded and answered here.
 *
 * The host also answers the worker's face acquisitions (`ooxmlAcquireFonts`):
 * after claiming a family, the worker session asks the page to fetch its
 * claimed http(s) / blob: faces. Only URLs this document's own resolution
 * returned for the named family are fetched (a request naming anything else
 * fails that URL), in this realm with `credentials: 'same-origin'` as CSS
 * font fetching uses, through the realm's shared acquisition cache
 * (`face-download.ts`) and the document's failure record, so the XLSX viewer
 * realm and the worker paint the same faces from one request. The worker's
 * remaining face-loading ceiling is the deadline; the reply carries a copy of
 * each body, the cache keeping the original until the document is disposed.
 */
import { fontSourceViolationFromWire } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import {
  FONT_DEADLINE_REASON,
  acquireResolverFacesInRealm,
  isDownloadableFace,
  releaseAcquiredFaces,
} from './face-download.js';
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import {
  copiedForTransfer,
  isWorkerFontResolveMessage,
  type BridgedFamilyAnswer,
  type FontAcquireOutcome,
  type FontResolveOutcome,
  type MainFontAcquireReply,
  type MainFontResolveReply,
  type WorkerFontAcquireRequest,
  type WorkerFontResolveMessage,
  type WorkerFontResolveRequest,
} from './resolver-bridge.js';
import type { FontResolverResponder } from './resolver-bridge-responder.js';
import {
  concludeBridgedRequest,
  decidedFamilyOutcome,
  settledBridgedRequest,
  type DocumentSourceCache,
  type FontResolverFailures,
} from './resolver-failures.js';
import { acquireSourceCache, type SourceCache } from './source-cache.js';
import type { FontSource } from './source.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

export interface FontResolverBridgeHostOptions {
  readonly post: Post;
  /** The document's validated `fonts` option. */
  readonly sources: readonly FontSource[] | undefined;
  /** The document's resolver failure record, when its viewer-realm sessions
   * share it (XLSX); see `resolver-resolution.ts`. */
  readonly resolverFailures?: FontResolverFailures;
  /** Loads the responder chunk; defaults to the lazy import (replaced only
   * by tests that exercise a chunk that fails to load). */
  readonly loadResponder?: () => Promise<typeof import('./resolver-bridge-responder.js')>;
}

export class FontResolverBridgeHost {
  private options: FontResolverBridgeHostOptions | null;
  private readonly postToWorker: Post;
  /** The document's outcome record, shared with the responder (and, for
   * XLSX, the viewer-realm sessions). */
  private readonly failures: FontResolverFailures;
  private responder: Promise<FontResolverResponder | null> | null = null;
  private disposed = false;
  /** This document's handle on each resolver source's realm cache, by its
   * index in `fonts`: the holder of every entry the host answers from. */
  private readonly caches = new Map<number, SourceCache>();
  private readonly cacheFor: DocumentSourceCache = (sourceIndex, source) => {
    let cache = this.caches.get(sourceIndex);
    if (!cache) {
      cache = acquireSourceCache(source);
      this.caches.set(sourceIndex, cache);
    }
    return cache;
  };
  /** Newest worker generation seen. */
  private generation = -1;
  /** Requests not yet handed to the responder, with their deadline. */
  private readonly waiting = new Map<string, {
    readonly request: WorkerFontResolveRequest;
    readonly deadline: number;
    readonly timer: ReturnType<typeof setTimeout>;
  }>();
  /** Face acquisitions in flight, by `generation:requestId`. */
  private readonly acquisitions = new Map<string, { readonly controller: AbortController; readonly generation: number }>();

  constructor(options: FontResolverBridgeHostOptions) {
    // Only a document configured with a resolver ever answers; any other
    // bridge message is consumed and ignored.
    this.failures = options.resolverFailures ?? new Map();
    this.postToWorker = options.post;
    this.options = Array.isArray(options.sources) && options.sources.some((source) => source.kind === 'font-resolver')
      ? options : null;
    if (this.options) void this.ensureResponder(this.options);
  }

  /** Route `message`; returns true when it belongs to the bridge. */
  accept(message: unknown): boolean {
    if (!isWorkerFontResolveMessage(message)) return false;
    const options = this.options;
    if (this.disposed) return true;
    const bridged = message as WorkerFontResolveMessage;
    if (!isNonNegativeSafeInteger(bridged.generation) || !isNonNegativeSafeInteger(bridged.requestId)) return true;
    if (!options?.sources) {
      // No resolver of this document can run here: the worker still gets an
      // answer (it never decides an outcome itself).
      if (bridged.kind === 'ooxmlResolveFonts') {
        this.post(bridged, { status: 'failed', reason: 'font resolver function is not available in this realm' });
      }
      return true;
    }
    const key = `${bridged.generation}:${bridged.requestId}`;
    if (bridged.kind === 'ooxmlAcquireFonts') {
      this.acquire(key, bridged, options.sources);
      return true;
    }
    if (bridged.kind === 'ooxmlCancelResolveFonts') {
      // An acquisition: aborted and never answered.
      const acquisition = this.acquisitions.get(key);
      if (acquisition) { acquisition.controller.abort(); this.acquisitions.delete(key); return true; }
      // Not handed to the responder yet: the call ends here, unanswered.
      const waiting = this.waiting.get(key);
      if (waiting) { clearTimeout(waiting.timer); this.waiting.delete(key); return true; }
    } else {
      if (bridged.generation < this.generation || this.waiting.has(key)) return true;
      if (bridged.generation > this.generation) {
        // The worker moved to a new document: its older calls are abandoned.
        this.generation = bridged.generation;
        for (const [waitingKey, waiting] of [...this.waiting]) {
          if (waiting.request.generation >= bridged.generation) continue;
          clearTimeout(waiting.timer);
          this.waiting.delete(waitingKey);
        }
        this.abandonAcquisitionsBefore(bridged.generation);
      }
      // Families this document already decided are answered now, whatever
      // the responder's state: the worker adopts the viewer realm's outcome.
      const split = settledBridgedRequest(this.failures, options.sources, bridged, this.cacheFor);
      if (split && split.unsettled.length === 0) {
        this.answer(bridged, split.settled);
        return true;
      }
      // The call's deadline starts on arrival: the worker's remaining ceiling
      // for its ensure, at most the font ceiling. Loading the responder counts
      // against it, and on expiry the call is failed here, recorded for the
      // document and answered, whether or not the responder has loaded.
      const { remainingMs } = bridged;
      const budget = typeof remainingMs === 'number' && Number.isFinite(remainingMs) && remainingMs >= 0
        ? Math.min(remainingMs, FONT_LOAD_CEILING_MS) : FONT_LOAD_CEILING_MS;
      const deadline = Date.now() + budget;
      const timer = setTimeout(() => {
        if (this.waiting.get(key)?.request !== bridged) return;
        this.waiting.delete(key);
        this.fail(bridged, 'timed out');
      }, budget);
      this.waiting.set(key, { request: bridged, deadline, timer });
    }
    const responder = this.ensureResponder(options);
    void responder.then((responder) => {
      if (this.disposed) return;
      if (bridged.kind === 'ooxmlCancelResolveFonts') { responder?.accept(bridged); return; }
      // Expired, cancelled or superseded while the responder loaded: not run.
      const waiting = this.waiting.get(key);
      if (waiting?.request !== bridged) return;
      clearTimeout(waiting.timer);
      this.waiting.delete(key);
      if (!responder) {
        // The responder could not load: fail the call instead of leaving the
        // worker waiting, with the same outcome for the whole document.
        this.fail(bridged, 'font resolver bridge unavailable');
        return;
      }
      // Only the time left of the call's deadline; an expired call is not run.
      const left = waiting.deadline - Date.now();
      if (left <= 0) { this.fail(bridged, 'timed out'); return; }
      const forwarded: WorkerFontResolveRequest = { ...bridged, remainingMs: left };
      responder.accept(forwarded);
    });
    return true;
  }

  private abandonAcquisitionsBefore(generation: number): void {
    for (const [key, acquisition] of [...this.acquisitions]) {
      if (acquisition.generation >= generation) continue;
      acquisition.controller.abort();
      this.acquisitions.delete(key);
    }
  }

  /** Answer one face acquisition (see the module comment). */
  private acquire(key: string, request: WorkerFontAcquireRequest, sources: readonly FontSource[]): void {
    if (request.generation < this.generation || this.acquisitions.has(key)) return;
    if (request.generation > this.generation) {
      this.generation = request.generation;
      this.abandonAcquisitionsBefore(request.generation);
    }
    const reply = (outcome: FontAcquireOutcome, transfer: ArrayBuffer[] = []) => {
      if (this.disposed) return;
      try {
        this.postToWorker({
          kind: 'ooxmlFontsAcquired', generation: request.generation, requestId: request.requestId, outcome,
        } satisfies MainFontAcquireReply, transfer);
      } catch { /* worker terminated */ }
    };
    const { sourceIndex, resolverId, remainingMs, families } = request;
    const source = typeof sourceIndex === 'number' && Number.isSafeInteger(sourceIndex) && sourceIndex >= 0
      ? sources[sourceIndex] : undefined;
    if (source?.kind !== 'font-resolver' || source.id !== resolverId || !Array.isArray(families)) {
      reply({ status: 'failed', reason: 'malformed font acquire request' });
      return;
    }
    // Only URLs of the faces this document's resolution returned for the
    // named family (its kept candidates), whatever the worker asked for.
    const cache = this.cacheFor(sourceIndex, source);
    const allowed = new Set<string>();
    const asked = new Set<string>();
    for (const entry of families as unknown[]) {
      const { family, urls } = (entry ?? {}) as { family?: unknown; urls?: unknown };
      if (typeof family !== 'string' || !Array.isArray(urls)) continue;
      for (const url of urls) if (typeof url === 'string') asked.add(url);
      const decided = decidedFamilyOutcome(this.failures, cache, family, cache);
      if (decided?.kind !== 'claimed') continue;
      for (const declaration of decided.declarations) if (isDownloadableFace(declaration)) allowed.add(declaration.url as string);
    }
    const budget = typeof remainingMs === 'number' && Number.isFinite(remainingMs) && remainingMs >= 0
      ? Math.min(remainingMs, FONT_LOAD_CEILING_MS) : FONT_LOAD_CEILING_MS;
    const controller = new AbortController();
    const acquisition = { controller, generation: request.generation };
    this.acquisitions.set(key, acquisition);
    const timer = setTimeout(() => controller.abort(FONT_DEADLINE_REASON), budget);
    const urls = [...asked].filter((url) => allowed.has(url));
    void acquireResolverFacesInRealm(this.failures, cache.id, urls, {
      signal: controller.signal, deadline: Date.now() + budget, holder: this,
    }).then((outcomes): [FontAcquireOutcome, ArrayBuffer[]] => {
      const transfer: ArrayBuffer[] = [];
      const faces = [...asked].map((url) => {
        const outcome = outcomes.get(url) ?? { reason: 'not a font this document resolved' };
        if ('reason' in outcome) return { url, reason: outcome.reason };
        // A copy: the realm cache keeps the original for the document.
        const data = outcome.data.slice(0);
        transfer.push(data);
        return { url, data };
      });
      return [{ status: 'faces', faces }, transfer];
    }, (error: unknown): [FontAcquireOutcome, ArrayBuffer[]] => {
      const violation = error instanceof OoxmlError ? fontSourceViolationFromWire(error.violation) : undefined;
      return [violation ? { status: 'limit', violation }
        : { status: 'failed', reason: error instanceof Error ? error.message : String(error) }, []];
    }).then(([outcome, transfer]) => {
      clearTimeout(timer);
      // Cancelled, superseded or disposed meanwhile: nothing waits for it.
      if (this.acquisitions.get(key) !== acquisition) return;
      this.acquisitions.delete(key);
      reply(outcome, transfer);
    });
  }

  /** Load the responder (once). Started eagerly for a document with a
   * resolver, so the bootstrap window is small; the deadline and settled
   * outcomes above do not depend on it. */
  private ensureResponder(options: FontResolverBridgeHostOptions): Promise<FontResolverResponder | null> {
    this.responder ??= (options.loadResponder ?? (() => import('./resolver-bridge-responder.js')))().then(({ FontResolverResponder: Responder }) => {
      if (this.disposed || !this.options?.sources) return null;
      return new Responder({
        post: options.post,
        sources: options.sources ?? [],
        resolverFailures: this.failures,
      });
    }, (error: unknown) => {
      console.warn(`[ooxml] font resolver bridge failed to load: ${error instanceof Error ? error.message : String(error)}`);
      return null;
    });
    return this.responder;
  }

  /** End a request the responder will not run (its deadline expired, or
   * the responder could not load), through the document's single
   * conclusion path (`concludeBridgedRequest`). */
  private fail(request: WorkerFontResolveRequest, reason: string): void {
    const options = this.options;
    if (this.disposed || !options?.sources) return;
    const outcome = concludeBridgedRequest(this.failures, options.sources, request, reason, this.cacheFor);
    if (outcome.status === 'families') this.answer(request, outcome.families);
    else this.post(request, outcome);
  }

  /** Answer per family; the worker receives copies of held buffers. */
  private answer(request: WorkerFontResolveRequest, families: readonly BridgedFamilyAnswer[]): void {
    const sent = copiedForTransfer(families);
    this.post(request, { status: 'families', families: sent.answers }, sent.transfer);
  }

  private post(request: WorkerFontResolveRequest, outcome: FontResolveOutcome, transfer: ArrayBuffer[] = []): void {
    if (this.disposed) return;
    try {
      this.postToWorker({
        kind: 'ooxmlFontsResolved', generation: request.generation, requestId: request.requestId, outcome,
      } satisfies MainFontResolveReply, transfer);
    } catch { /* worker terminated */ }
  }

  /** The document is destroyed: abort every in-flight resolver call. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.options = null;
    for (const { timer } of this.waiting.values()) clearTimeout(timer);
    this.waiting.clear();
    for (const { controller } of this.acquisitions.values()) controller.abort();
    this.acquisitions.clear();
    releaseAcquiredFaces(this);
    // Every entry this document answered from is released with it.
    for (const cache of this.caches.values()) cache.release();
    this.caches.clear();
    void this.responder?.then((responder) => responder?.dispose());
  }
}
