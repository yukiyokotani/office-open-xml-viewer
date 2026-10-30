/**
 * `fontResolver` worker bridge: main-thread half (`resolver-bridge.ts` has the
 * protocol). Loaded lazily, on a document's first bridged request, by
 * {@link FontResolverBridgeHost} (`resolver-bridge-host.ts`).
 *
 * Library API policy, not OOXML behaviour. For each `ooxmlResolveFonts`
 * request the responder looks the resolver id up among the document's own
 * `fonts` descriptors, so a worker can only run functions this document was
 * configured with, and resolves the demanded families through this realm's
 * resolution cache (`resolver-resolution.ts`, design §4.4): a family another
 * requester of this realm already resolved or is resolving (an XLSX viewer
 * realm session, another document sharing the resolver) is joined, not asked
 * again, so the application function runs once per family in both modes. A
 * new call receives a copy of the demand (family names only) and an
 * `AbortSignal` that the worker's cancellation, document destroy and the 15 s
 * font ceiling abort (a shared call is aborted only when every requester left).
 * The reply carries each family's candidate declarations (the worker claims
 * from them under the scope its own text proves), failure or absence, or the
 * breached ceiling. The host downloads the candidates' URL faces in this
 * realm before posting (`resolver-bridge-host.ts`).
 *
 * Holding and bytes (design §4.3, §4.4): the bridge holds every entry it
 * joined through the document's cache handle until the document is destroyed,
 * exactly as a main-mode session holds its entries, so live documents share
 * one resolution per family whichever realm paints them (a second worker-mode
 * document, the XLSX viewer realm). Main's cache therefore keeps the bytes the
 * worker receives, and a buffer the cache keeps must never be detached: the
 * reply carries one copy per distinct buffer, made while the entries are held.
 * The cache charges what main keeps and the worker's session what the worker
 * keeps, each once; `dispose` returns every entry. A failure is not kept by
 * the realm cache, but it is recorded in the document's failure record
 * (`resolverFailures`, shared with the XLSX viewer sessions), so one document
 * sees one outcome per family. Requests are named by the source's index in the
 * document's `fonts`, validated against that descriptor's id, so two
 * configurations of one function stay apart. A reply for a request the worker
 * cancelled (or abandoned with a newer generation), or produced after the
 * document was destroyed, is never posted.
 */
import type { OoxmlError } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_SOURCE_FAMILIES } from '../worker/resource-policy.generated.js';
import { FONT_LOAD_CEILING_MS } from './font-registry.js';
import {
  copiedForTransfer,
  type BridgedFamilyAnswer,
  type FontResolveOutcome,
  type MainFontResolveReply,
  type WorkerFontResolveRequest,
} from './resolver-bridge.js';
import { concludeBridgedRequest } from './resolver-failures.js';
import {
  resolveFamiliesThroughCache,
  runResolverInRealm,
  type FontResolverFailures,
} from './resolver-resolution.js';
import { acquireSourceCache, type SourceCache } from './source-cache.js';
import {
  fontResolverFunction,
  type FontDemand,
  type FontResolverSource,
  type FontSource,
} from './source.js';

type Post = (message: unknown, transfer?: Transferable[]) => void;

export interface FontResolverResponderOptions {
  /** Posts to the document's render worker. */
  readonly post: Post;
  /** The document's validated `fonts` option (main-realm descriptors). */
  readonly sources: readonly FontSource[];
  /** The document's resolver failure record, shared with its viewer-realm
   * sessions (XLSX); a private one otherwise. */
  readonly resolverFailures?: FontResolverFailures;
}

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;

function names(value: unknown): readonly string[] | undefined {
  if (!Array.isArray(value) || value.length > HARD_MAX_FONT_SOURCE_FAMILIES) return undefined;
  return value.every((name) => typeof name === 'string') ? Object.freeze([...value as string[]]) : undefined;
}

/**
 * The application's copy of a bridged demand, or undefined when the worker
 * sent something no library session produces: exactly `families` and
 * `fallbackFamilies`, string lists within `HARD_MAX_FONT_SOURCE_FAMILIES`
 * together.
 */
export function copyBridgedFontDemand(value: unknown): FontDemand | undefined {
  if (value === null || typeof value !== 'object') return undefined;
  const fields = value as Partial<Record<keyof FontDemand, unknown>>;
  if (Object.keys(fields).some((key) => key !== 'families' && key !== 'fallbackFamilies')) return undefined;
  const families = names(fields.families);
  const fallbackFamilies = names(fields.fallbackFamilies);
  if (!families || !fallbackFamilies || families.length + fallbackFamilies.length > HARD_MAX_FONT_SOURCE_FAMILIES) {
    return undefined;
  }
  return Object.freeze({ families, fallbackFamilies });
}

function isFontSourceLimit(error: unknown): error is OoxmlError & { violation: NonNullable<OoxmlError['violation']> } {
  return error instanceof Error && (error as OoxmlError).code === 'font-source-limit'
    && (error as OoxmlError).violation !== undefined;
}

/** One document's main-side bridge. */
export class FontResolverResponder {
  private readonly post: Post;
  private sources: readonly FontSource[];
  private readonly failures: FontResolverFailures;
  private disposed = false;
  /** Newest worker generation seen; older calls are abandoned. */
  private generation = -1;
  /** In-flight calls by `${generation}:${requestId}` (library-generated keys). */
  private readonly calls = new Map<string, { readonly controller: AbortController; readonly generation: number }>();
  /** This document's hold on each resolver source's realm cache, by its
   * index in the document's `fonts`; every joined entry is held by it. */
  private readonly caches = new Map<number, SourceCache>();

  constructor(options: FontResolverResponderOptions) {
    this.post = options.post;
    this.sources = options.sources;
    this.failures = options.resolverFailures ?? new Map();
  }

  /** Handle one worker → main bridge message (already type-guarded). */
  accept(message: { readonly kind: string }): void {
    if (this.disposed) return;
    const fields = message as Partial<WorkerFontResolveRequest>;
    if (!isNonNegativeSafeInteger(fields.generation) || !isNonNegativeSafeInteger(fields.requestId)) return;
    const key = `${fields.generation}:${fields.requestId}`;
    if (message.kind === 'ooxmlCancelResolveFonts') {
      this.calls.get(key)?.controller.abort();
      this.calls.delete(key);
      return;
    }
    if (message.kind !== 'ooxmlResolveFonts' || this.calls.has(key)) return;
    if (fields.generation < this.generation) return;
    if (fields.generation > this.generation) {
      // The worker moved to a new document: whatever it asked before is
      // abandoned (its own reset also cancels each call).
      this.generation = fields.generation;
      for (const [callKey, call] of [...this.calls]) {
        if (call.generation >= fields.generation) continue;
        call.controller.abort();
        this.calls.delete(callKey);
      }
    }
    void this.respond(key, fields.generation, fields.requestId, fields);
  }

  /** The document is destroyed: abort every call and return every cache
   * entry this document held; nothing is posted after. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    const calls = [...this.calls.values()];
    this.calls.clear();
    for (const { controller } of calls) controller.abort();
    for (const cache of this.caches.values()) cache.release();
    this.caches.clear();
    // The descriptors may hold application font bytes.
    this.sources = [];
    this.failures.clear();
  }

  /** In-flight calls (tests and leak checks). */
  get pendingCount(): number { return this.calls.size; }

  private async respond(
    key: string, generation: number, requestId: number, request: Partial<WorkerFontResolveRequest>,
  ): Promise<void> {
    const controller = new AbortController();
    const call = { controller, generation };
    this.calls.set(key, call);
    // The single path for a failed reply (`concludeBridgedRequest`): families
    // the document already decided adopt that outcome, the rest are recorded.
    const fail = (reason: string): FontResolveOutcome => (this.disposed
      // Nothing is posted after dispose; never take a hold for a gone document.
      ? { status: 'failed', reason }
      : concludeBridgedRequest(this.failures, this.sources, request, reason, (index, source) => this.cacheOf(index, source)));
    // The worker's remaining ceiling for its current ensure: main mode's
    // deadline is cumulative over the source walk, and so is this one.
    const { remainingMs } = request;
    const budget = typeof remainingMs === 'number' && Number.isFinite(remainingMs) && remainingMs >= 0
      ? Math.min(remainingMs, FONT_LOAD_CEILING_MS) : FONT_LOAD_CEILING_MS;
    let expired = false;
    const timer = setTimeout(() => { expired = true; controller.abort(); }, budget);
    let outcome: FontResolveOutcome;
    try {
      outcome = await this.resolve(request, controller.signal, () => expired, fail);
    } finally {
      clearTimeout(timer);
    }
    // Cancelled by the worker or the document destroyed: the worker no longer
    // waits, so the answer is dropped here (cached entries stay with their
    // holders and are released with them).
    const live = this.calls.get(key) === call;
    this.calls.delete(key);
    if (this.disposed || !live || (controller.signal.aborted && !expired)) return;
    // Expired: the document's conclusion (idempotent; a family decided by
    // another requester meanwhile keeps that outcome).
    if (expired) outcome = fail('timed out');
    let reply = outcome;
    let transfer: ArrayBuffer[] = [];
    if (outcome.status === 'families') {
      // Copied while this document's hold keeps the originals cached.
      const sent = copiedForTransfer(outcome.families);
      reply = { status: 'families', families: sent.answers };
      transfer = sent.transfer;
    }
    try {
      this.post({ kind: 'ooxmlFontsResolved', generation, requestId, outcome: reply } satisfies MainFontResolveReply, transfer);
    } catch (error) {
      // The worker is gone (nothing waits), or the reply did not clone: fail
      // the call's families instead of leaving the worker waiting.
      const failed = fail(error instanceof Error ? error.message : String(error));
      try {
        this.post({ kind: 'ooxmlFontsResolved', generation, requestId, outcome: failed } satisfies MainFontResolveReply);
      } catch { /* worker terminated */ }
    }
  }

  private cacheOf(sourceIndex: number, source: FontResolverSource): SourceCache {
    let cache = this.caches.get(sourceIndex);
    if (!cache) {
      cache = acquireSourceCache(source);
      this.caches.set(sourceIndex, cache);
    }
    return cache;
  }

  private async resolve(
    request: Partial<WorkerFontResolveRequest>,
    signal: AbortSignal,
    expired: () => boolean,
    fail: (reason: string) => FontResolveOutcome,
  ): Promise<FontResolveOutcome> {
    const { sourceIndex, resolverId } = request;
    const demand = copyBridgedFontDemand(request.demand);
    if (!demand || typeof resolverId !== 'string' || !isNonNegativeSafeInteger(sourceIndex)) {
      return fail('malformed font resolver request');
    }
    // Exactly the document's own `fonts[sourceIndex]`: the id names the
    // function, the index the configuration (its substitutes).
    const source = this.sources[sourceIndex];
    if (source?.kind !== 'font-resolver' || source.id !== resolverId || !fontResolverFunction(source)) {
      return fail('font resolver function is not available in this realm');
    }
    const families = [
      ...demand.families.map((name) => ({ name, fallback: false })),
      ...demand.fallbackFamilies.map((name) => ({ name, fallback: true })),
    ];
    const cache = this.cacheOf(sourceIndex, source);
    const started = Date.now();
    try {
      const answers = await resolveFamiliesThroughCache({
        source, cache, runner: runResolverInRealm, families, signal,
        // Held until dispose, like a main-mode session's entries.
        holder: cache, sourceIndex, failures: this.failures,
        remainingMs: () => Math.max(0, (request.remainingMs ?? FONT_LOAD_CEILING_MS) - (Date.now() - started)),
      });
      // Main owns the deadline of a bridged call (the worker session waits
      // for this reply): a timeout is the document's outcome for these
      // families, so a viewer-realm session of the same document reports it.
      if (expired()) return fail('timed out');
      return {
        status: 'families',
        families: [...answers].map(([family, answer]): BridgedFamilyAnswer =>
          answer.kind === 'claimed' ? { family, kind: 'claimed', declarations: answer.declarations }
            : answer.kind === 'failed' ? { family, kind: 'failed', reason: answer.reason }
              : { family, kind: 'absent' }),
      };
    } catch (error) {
      if (expired()) return fail('timed out');
      if (isFontSourceLimit(error)) {
        // Concluded per family like any other batch end; the breach itself
        // still rejects the worker's load.
        fail(error.message);
        return { status: 'limit', violation: error.violation };
      }
      return fail(error instanceof Error ? error.message : String(error));
    }
  }
}
