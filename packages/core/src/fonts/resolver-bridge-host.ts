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
 * Every reply that carries faces leaves through {@link FontResolverBridgeHost}
 * `deliver`: the page downloads each http(s) / blob: `url` face the reply
 * names (`face-download.ts`, `credentials: 'same-origin'` as CSS font fetching
 * uses) and the worker receives the bytes, so a resolver's URLs are fetched by
 * the page in both modes and the worker fetches none of them. The downloads
 * run inside the call's deadline. A face whose download fails is left out of
 * the reply and its family warned once per document; a family left with no
 * face fails with the download's reason; a body beyond the face ceiling
 * rejects the worker's load as any `font-source-limit` breach does.
 */
import { fontSourceViolationFromWire } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import { downloadFontFaces, isDownloadableFace } from './face-download.js';
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import {
  copiedForTransfer,
  isWorkerFontResolveMessage,
  type BridgedFamilyAnswer,
  type FontResolveOutcome,
  type MainFontResolveReply,
  type WorkerFontResolveMessage,
  type WorkerFontResolveRequest,
} from './resolver-bridge.js';
import type { FontResolverResponder } from './resolver-bridge-responder.js';
import {
  concludeBridgedRequest,
  settledBridgedRequest,
  type DocumentSourceCache,
  type FontResolverFailures,
} from './resolver-failures.js';
import { acquireSourceCache, type SourceCache } from './source-cache.js';
import type { FontSource } from './source.js';
import { fontFamilyKey } from './substitutions.js';

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
  /** Deadline of every request not yet answered, by `generation:requestId`
   * (the downloads of its reply end there). */
  private readonly deadlines = new Map<string, number>();
  /** Face downloads in flight, by request key (a cancel aborts them). */
  private readonly downloads = new Map<string, AbortController>();
  /** Families already warned about a failed face download. */
  private readonly warned = new Set<string>();

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
    if (bridged.kind === 'ooxmlCancelResolveFonts') {
      // Downloading its reply: the worker no longer waits for it.
      this.downloads.get(key)?.abort();
      this.downloads.delete(key);
      this.deadlines.delete(key);
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
        const older = (entryKey: string) => Number(entryKey.slice(0, entryKey.indexOf(':'))) < bridged.generation;
        for (const entryKey of [...this.deadlines.keys()]) if (older(entryKey)) this.deadlines.delete(entryKey);
        for (const [entryKey, controller] of [...this.downloads]) {
          if (!older(entryKey)) continue;
          controller.abort();
          this.downloads.delete(entryKey);
        }
      }
      // The call's deadline starts on arrival: the worker's remaining ceiling
      // for its ensure, at most the font ceiling. Loading the responder and
      // downloading the reply's faces count against it, and on expiry the
      // call is failed here, recorded for the document and answered, whether
      // or not the responder has loaded.
      const { remainingMs } = bridged;
      const budget = typeof remainingMs === 'number' && Number.isFinite(remainingMs) && remainingMs >= 0
        ? Math.min(remainingMs, FONT_LOAD_CEILING_MS) : FONT_LOAD_CEILING_MS;
      const deadline = Date.now() + budget;
      this.deadlines.set(key, deadline);
      // Families this document already decided are answered now, whatever
      // the responder's state: the worker adopts the viewer realm's outcome.
      const split = settledBridgedRequest(this.failures, options.sources, bridged, this.cacheFor);
      if (split && split.unsettled.length === 0) {
        this.answer(bridged, split.settled);
        return true;
      }
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

  /** Load the responder (once). Started eagerly for a document with a
   * resolver, so the bootstrap window is small; the deadline and settled
   * outcomes above do not depend on it. */
  private ensureResponder(options: FontResolverBridgeHostOptions): Promise<FontResolverResponder | null> {
    this.responder ??= (options.loadResponder ?? (() => import('./resolver-bridge-responder.js')))().then(({ FontResolverResponder: Responder }) => {
      if (this.disposed || !this.options?.sources) return null;
      return new Responder({
        post: (message, transfer) => this.deliver(message as MainFontResolveReply, transfer as ArrayBuffer[] | undefined),
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
    this.deliver({
      kind: 'ooxmlFontsResolved', generation: request.generation, requestId: request.requestId,
      outcome: { status: 'families', families: sent.answers },
    }, sent.transfer);
  }

  private post(request: WorkerFontResolveRequest, outcome: FontResolveOutcome, transfer: ArrayBuffer[] = []): void {
    this.deliver({
      kind: 'ooxmlFontsResolved', generation: request.generation, requestId: request.requestId, outcome,
    }, transfer);
  }

  /**
   * The single exit of a reply (the host's own and the responder's). A reply
   * without URL faces is posted at once; otherwise the page downloads them
   * first, within the request's deadline, and the reply carries the bytes.
   * A post that fails (the reply did not clone) fails the request's families
   * instead of leaving the worker waiting.
   */
  private deliver(reply: MainFontResolveReply, transfer: ArrayBuffer[] = []): void {
    if (this.disposed) return;
    const key = `${reply.generation}:${reply.requestId}`;
    const { outcome } = reply;
    const urls = outcome.status === 'families'
      && outcome.families.some((answer) => answer.kind === 'claimed' && answer.declarations.some(isDownloadableFace));
    if (!urls || outcome.status !== 'families') {
      this.deadlines.delete(key);
      this.postNow(reply, transfer);
      return;
    }
    const controller = new AbortController();
    this.downloads.get(key)?.abort();
    this.downloads.set(key, controller);
    const deadline = this.deadlines.get(key) ?? Date.now() + FONT_LOAD_CEILING_MS;
    this.deadlines.delete(key);
    const timer = setTimeout(() => controller.abort(), Math.max(0, deadline - Date.now()));
    const faces = outcome.families.flatMap((answer) => (answer.kind === 'claimed' ? answer.declarations : []));
    void downloadFontFaces(faces, 'same-origin', controller.signal).then((downloads) => {
      const families = outcome.families.map((answer): BridgedFamilyAnswer => {
        if (answer.kind !== 'claimed') return answer;
        const kept = answer.declarations.flatMap((declaration) => downloads.replace(declaration) ?? []);
        const lost = answer.declarations.filter((declaration) => downloads.replace(declaration) === undefined);
        if (!lost.length) return { ...answer, declarations: kept };
        const reason = downloads.failures.get(lost[0].url as string) ?? 'font download failed';
        if (!kept.length) return { family: answer.family, kind: 'failed', reason };
        this.warnDownload(lost[0].family, reason);
        return { ...answer, declarations: kept };
      });
      return { outcome: { status: 'families', families } satisfies FontResolveOutcome, buffers: downloads.buffers };
    }, (error: unknown) => {
      const violation = error instanceof OoxmlError ? fontSourceViolationFromWire(error.violation) : undefined;
      const failed: FontResolveOutcome = violation
        ? { status: 'limit', violation }
        : { status: 'failed', reason: error instanceof Error ? error.message : String(error) };
      return { outcome: failed, buffers: [] as ArrayBuffer[] };
    }).then(({ outcome: sent, buffers }) => {
      clearTimeout(timer);
      // Cancelled (or superseded) meanwhile: the worker no longer waits.
      if (this.downloads.get(key) !== controller) return;
      this.downloads.delete(key);
      if (this.disposed) return;
      this.postNow({ ...reply, outcome: sent }, [...transfer, ...buffers]);
    });
  }

  private postNow(reply: MainFontResolveReply, transfer: ArrayBuffer[]): void {
    if (this.disposed) return;
    try {
      this.postToWorker(reply, transfer);
    } catch (error) {
      if (reply.outcome.status !== 'families') return;
      try {
        this.postToWorker({
          ...reply,
          outcome: { status: 'failed', reason: error instanceof Error ? error.message : String(error) },
        } satisfies MainFontResolveReply);
      } catch { /* worker terminated */ }
    }
  }

  private warnDownload(family: string, reason: string): void {
    if (this.warned.has(fontFamilyKey(family))) return;
    this.warned.add(fontFamilyKey(family));
    console.warn(`[ooxml] failed to preload web font(s): ${family} (${reason}); falling back to system fonts (text may shift or differ).`);
  }

  /** The document is destroyed: abort every in-flight resolver call. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.options = null;
    for (const { timer } of this.waiting.values()) clearTimeout(timer);
    this.waiting.clear();
    for (const controller of this.downloads.values()) controller.abort();
    this.downloads.clear();
    this.deadlines.clear();
    // Every entry this document answered from is released with it.
    for (const cache of this.caches.values()) cache.release();
    this.caches.clear();
    void this.responder?.then((responder) => responder?.dispose());
  }
}
