/**
 * `fontResolver` worker bridge: protocol and render-worker half.
 *
 * Library API policy, not OOXML behaviour. A {@link fontResolver} descriptor
 * carries only an id; its application function stays in the main realm that
 * called the factory (`source.ts`), so credentials, tokens and private
 * network clients never reach the worker (design §3.4). When the render worker
 * paints, its {@link FontSourceSession} runs every resolver call through
 * {@link WorkerFontResolverClient.run}: the call is posted to main
 * (`ooxmlResolveFonts`, naming the source by its index in the document's
 * `fonts`), main runs the function with an `AbortSignal` and replies
 * (`ooxmlFontsResolved`) with a copy of every `data` buffer in the transfer
 * list (`resolver-bridge-responder.ts`). Main resolves through its own realm
 * resolution cache, so the function runs once per family however many
 * realms ask (`resolver-resolution.ts`), and answers per family with the
 * family's candidate faces. The request carries only the demanded family
 * names (`FontDemand`): painted scalars and substitute scopes stay in the
 * worker, whose session admits each family's candidates and claims from them
 * under its own scope, exactly as it claims a direct answer.
 *
 * The worker never fetches a resolver URL. Once it has claimed a family, its
 * session asks main for the claimed faces' http(s) / blob: URLs through
 * {@link WorkerFontResolverClient.acquire} (`ooxmlAcquireFonts`, within the
 * worker's face-loading ceiling); main fetches only URLs its own resolution
 * of this document returned and replies (`ooxmlFontsAcquired`) with a copy of
 * each body, so cookies, HTTP authentication and Referer checks see the page
 * in both modes (`resolver-bridge-host.ts`, `face-download.ts`).
 *
 * Correlation follows the host-font measurement requests of the DOCX render
 * worker: each request carries the client's generation and a request id. A
 * worker that starts a new document calls {@link WorkerFontResolverClient.reset},
 * which abandons every pending call and posts a cancellation for each; main
 * also abandons the calls of a generation once a newer one asks. A reply for
 * an abandoned call, an older generation or an unknown id is dropped. When the
 * worker session abandons a call (its 15 s ceiling, release, every waiter
 * gone) the client likewise posts `ooxmlCancelResolveFonts`, which aborts the
 * application function's signal on main; main never posts for it.
 *
 * Replies are validated fail-closed: anything but per-family answers within
 * `HARD_MAX_FONT_SOURCE_FAMILIES` (claimed declarations within
 * `HARD_MAX_FONT_SOURCE_DECLARATIONS`, a failure reason, absence), a failure
 * reason or a well-formed limit violation fails the call's families (host fallback), and the
 * declarations themselves are validated again by the session.
 */
import { createAbortError } from '../internal/abort.js';
import { fontSourceViolationFromWire } from '../errors/font-source-limit.js';
import type { OoxmlFontSourceViolation } from '../errors/ooxml-error.js';
import {
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import type { FaceAcquisition, ResolverFaceRequest } from './face-download.js';
import { bridgedFontSourceLimitError } from './resolver-admission.js';
import { BridgedFamilyResolutions, type ResolverCallInfo } from './resolver-resolution.js';
import {
  validateFontFaceDeclaration,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolveContext,
  type FontResolverSource,
} from './source.js';

/** Worker → main: run the document's `fonts[sourceIndex]` (a resolver whose
 * id is `resolverId`) for `demand`. */
export interface WorkerFontResolveRequest {
  readonly kind: 'ooxmlResolveFonts';
  readonly generation: number;
  readonly requestId: number;
  readonly sourceIndex: number;
  readonly resolverId: string;
  /** Milliseconds left of the worker session's font ceiling for the ensure
   * that asks; main's deadline for the call (see ResolverCallInfo). */
  readonly remainingMs: number;
  readonly demand: FontDemand;
}

/** Worker → main: fetch the claimed URL faces of the document's
 * `fonts[sourceIndex]` (requests share the resolve request id space). */
export interface WorkerFontAcquireRequest {
  readonly kind: 'ooxmlAcquireFonts';
  readonly generation: number;
  readonly requestId: number;
  readonly sourceIndex: number;
  readonly resolverId: string;
  /** Milliseconds left of the worker session's face-loading ceiling. */
  readonly remainingMs: number;
  readonly families: readonly { readonly family: string; readonly urls: readonly string[] }[];
}

/** Main's outcome for one acquire request. */
export type FontAcquireOutcome =
  /** Per URL: the body (in the transfer list) or why it failed. */
  | { readonly status: 'faces'; readonly faces: readonly ({ readonly url: string; readonly data: ArrayBuffer }
    | { readonly url: string; readonly reason: string })[] }
  | { readonly status: 'failed'; readonly reason: string }
  | { readonly status: 'limit'; readonly violation: OoxmlFontSourceViolation };

/** Main → worker. */
export interface MainFontAcquireReply {
  readonly kind: 'ooxmlFontsAcquired';
  readonly generation: number;
  readonly requestId: number;
  readonly outcome: FontAcquireOutcome;
}

/** Worker → main: the worker no longer waits for `requestId`. */
export interface WorkerFontResolveCancel {
  readonly kind: 'ooxmlCancelResolveFonts';
  readonly generation: number;
  readonly requestId: number;
}

/** Main's answer for one demanded family (lower-cased `family`). */
export type BridgedFamilyAnswer =
  /** Its candidate declarations (`resolverCandidates`); every `data` buffer
   * is in the transfer list, and no declaration is an http(s) or blob: URL. */
  | { readonly family: string; readonly kind: 'claimed'; readonly declarations: readonly FontFaceDeclaration[] }
  | { readonly family: string; readonly kind: 'failed'; readonly reason: string }
  | { readonly family: string; readonly kind: 'absent' };

/** What main answers for one request. */
export type FontResolveOutcome =
  /** Per-family answers from main's resolution cache. */
  | { readonly status: 'families'; readonly families: readonly BridgedFamilyAnswer[] }
  /** The request could not be served (host fallback for every family). */
  | { readonly status: 'failed'; readonly reason: string }
  /** The answer breached a font-source ceiling (rejects the load). */
  | { readonly status: 'limit'; readonly violation: OoxmlFontSourceViolation };

/** Main → worker. */
export interface MainFontResolveReply {
  readonly kind: 'ooxmlFontsResolved';
  readonly generation: number;
  readonly requestId: number;
  readonly outcome: FontResolveOutcome;
}

export type WorkerFontResolveMessage = WorkerFontResolveRequest | WorkerFontAcquireRequest | WorkerFontResolveCancel;

type Post = (message: unknown, transfer?: Transferable[]) => void;

const kindOf = (value: unknown): unknown =>
  value !== null && typeof value === 'object' ? (value as { kind?: unknown }).kind : undefined;

/** Whether `value` is a worker → main resolver-bridge message (main side). */
export function isWorkerFontResolveMessage(value: unknown): value is WorkerFontResolveMessage {
  const kind = kindOf(value);
  return kind === 'ooxmlResolveFonts' || kind === 'ooxmlAcquireFonts' || kind === 'ooxmlCancelResolveFonts';
}

/** Whether `value` is a main → worker resolver-bridge reply (worker side). */
export function isMainFontResolveReply(value: unknown): value is MainFontResolveReply | MainFontAcquireReply {
  const kind = kindOf(value);
  return kind === 'ooxmlFontsResolved' || kind === 'ooxmlFontsAcquired';
}

/** The reply's answers with one copy per distinct buffer: the originals are
 * kept by main's cache (and possibly the application), so none is detached. */
export function copiedForTransfer(
  answers: readonly BridgedFamilyAnswer[],
): { answers: readonly BridgedFamilyAnswer[]; transfer: ArrayBuffer[] } {
  const copies = new Map<ArrayBuffer, ArrayBuffer>();
  const wire = (declaration: FontFaceDeclaration): FontFaceDeclaration => {
    const { data } = declaration;
    if (!data) return declaration;
    let copy = copies.get(data);
    if (!copy) {
      copy = data.slice(0);
      copies.set(data, copy);
    }
    return { ...declaration, data: copy };
  };
  const out = answers.map((answer) => (answer.kind === 'claimed'
    ? { ...answer, declarations: answer.declarations.map(wire) } : answer));
  return { answers: out, transfer: [...copies.values()] };
}

/** Longest failure reason kept from a reply (it only reaches console.warn). */
const MAX_REASON_LENGTH = 1024;

interface PendingCall {
  readonly kind: 'ooxmlFontsResolved' | 'ooxmlFontsAcquired';
  readonly resolve: (answer: never) => void;
  readonly reject: (error: unknown) => void;
  readonly detach: () => void;
  readonly generation: number;
}

/**
 * Render-worker half of the bridge: a `FontResolverRunner` for the worker's
 * font-source session. One client per worker.
 */
export class WorkerFontResolverClient {
  private generation = 0;
  private nextRequestId = 1;
  private readonly pending = new Map<number, PendingCall>();

  constructor(private readonly post: Post) {}

  /** Spread into `FontSourceSessionOptions`: the runner, and the fact that
   * main owns the deadline of a bridged call (it times the call out and
   * records the outcome for the whole document, so a workbook's viewer realm
   * and worker agree), so the session does not abandon the wait itself. */
  get sessionOptions(): {
    readonly runResolver: WorkerFontResolverClient['run'];
    readonly resolverCeilingOwnedByRunner: true;
    readonly acquireResolverFaces: WorkerFontResolverClient['acquire'];
    readonly acquiredFacesOwnedBySession: true;
  } {
    return {
      runResolver: this.run, resolverCeilingOwnedByRunner: true,
      acquireResolverFaces: this.acquire, acquiredFacesOwnedBySession: true,
    };
  }

  /**
   * A `ResolverFaceAcquirer` (`source-session.ts`): main fetches the claimed
   * URL faces. The session's face-loading ceiling bounds the wait: when its
   * signal aborts, the call is cancelled on main and every face not yet
   * answered fails with `'timed out'`.
   */
  readonly acquire = (
    request: ResolverFaceRequest,
    { signal, deadline }: { readonly signal: AbortSignal; readonly deadline: number },
  ): Promise<ReadonlyMap<string, FaceAcquisition>> => {
    const urls = [...new Set(request.families.flatMap(({ urls: list }) => list))];
    const timedOut = () => new Map<string, FaceAcquisition>(urls.map((url) => [url, { reason: 'timed out' }]));
    if (signal.aborted) return Promise.resolve(timedOut());
    const generation = this.generation;
    const requestId = this.nextRequestId++;
    return new Promise<ReadonlyMap<string, FaceAcquisition>>((resolve, reject) => {
      const onAbort = () => {
        if (!this.settle(requestId)) return;
        resolve(timedOut());
        this.postQuietly({ kind: 'ooxmlCancelResolveFonts', generation, requestId });
      };
      signal.addEventListener('abort', onAbort, { once: true });
      this.pending.set(requestId, {
        kind: 'ooxmlFontsAcquired',
        resolve: ((outcome: FontAcquireOutcome) => {
          if (outcome.status === 'limit') { reject(bridgedFontSourceLimitError(outcome.violation)); return; }
          const answers = new Map<string, FaceAcquisition>(urls.map((url) => [url, {
            reason: outcome.status === 'failed' ? outcome.reason : 'font download failed',
          }]));
          if (outcome.status === 'faces') {
            for (const face of outcome.faces) if (answers.has(face.url)) answers.set(face.url, 'data' in face ? { data: face.data } : { reason: face.reason });
          }
          resolve(answers);
        }) as (answer: never) => void,
        // A reset abandons the document: nothing waits for these faces.
        reject: () => resolve(timedOut()),
        generation,
        detach: () => signal.removeEventListener('abort', onAbort),
      });
      try {
        this.post({
          kind: 'ooxmlAcquireFonts', generation, requestId, sourceIndex: request.sourceIndex,
          resolverId: request.source.id, remainingMs: Math.max(0, deadline - Date.now()),
          families: request.families,
        } satisfies WorkerFontAcquireRequest);
      } catch {
        this.settle(requestId);
        resolve(timedOut());
      }
    });
  };

  /** Pass as `FontSourceSessionOptions.runResolver`. */
  readonly run = (
    source: FontResolverSource,
    demand: FontDemand,
    context: FontResolveContext,
    { sourceIndex, remainingMs }: ResolverCallInfo,
  ): Promise<BridgedFamilyResolutions> => {
    const { signal } = context;
    if (signal.aborted) return Promise.reject(createAbortError('Font resolver abandoned'));
    const generation = this.generation;
    const requestId = this.nextRequestId++;
    return new Promise<BridgedFamilyResolutions>((resolve, reject) => {
      const abandon = (error: Error) => {
        if (!this.settle(requestId)) return;
        reject(error);
        this.postQuietly({ kind: 'ooxmlCancelResolveFonts', generation, requestId } satisfies WorkerFontResolveCancel);
      };
      const onAbort = () => abandon(createAbortError('Font resolver abandoned'));
      signal.addEventListener('abort', onAbort, { once: true });
      // Main is the single authority on the call's outcome (it enforces the
      // deadline and records the outcome for the whole document), so the
      // worker never decides one itself: it waits for main's reply while its
      // document lives. Every main-side path ends in a reply, or in the
      // document's destruction, which resets or terminates this worker.
      this.pending.set(requestId, {
        kind: 'ooxmlFontsResolved', resolve: resolve as (answer: never) => void, reject, generation,
        detach: () => signal.removeEventListener('abort', onAbort),
      });
      try {
        this.post({
          kind: 'ooxmlResolveFonts', generation, requestId, sourceIndex, resolverId: source.id, demand, remainingMs,
        } satisfies WorkerFontResolveRequest);
      } catch (error) {
        this.settle(requestId);
        reject(error);
      }
    });
  };

  /** Route a reply; returns true when `message` belongs to this bridge (it is
   * consumed even when stale, so its bytes are dropped here). */
  accept(message: unknown): boolean {
    if (!isMainFontResolveReply(message)) return false;
    const { generation, requestId } = message as Partial<MainFontResolveReply>;
    if (generation !== this.generation || typeof requestId !== 'number') return true;
    const call = this.pending.get(requestId);
    if (!call || call.kind !== message.kind || !this.settle(requestId)) return true;
    if (message.kind === 'ooxmlFontsAcquired') {
      (call.resolve as (outcome: FontAcquireOutcome) => void)(readAcquireOutcome(message.outcome));
      return true;
    }
    const outcome = readOutcome(message.outcome);
    if (outcome.status === 'families') {
      (call.resolve as (answers: BridgedFamilyResolutions) => void)(
        new BridgedFamilyResolutions(new Map(outcome.families.map((answer) => [answer.family, answer]))));
    }
    else if (outcome.status === 'limit') call.reject(bridgedFontSourceLimitError(outcome.violation));
    else call.reject(new Error(outcome.reason));
    return true;
  }

  /** A new document starts (or the worker's document is gone): abandon every
   * pending call, cancel it on main, and ignore replies of the previous
   * generation. */
  reset(): void {
    this.generation++;
    const calls = [...this.pending];
    this.pending.clear();
    for (const [requestId, call] of calls) {
      call.detach();
      call.reject(createAbortError('Font resolver abandoned'));
      // Main aborts the application call and posts nothing for it.
      this.postQuietly({ kind: 'ooxmlCancelResolveFonts', generation: call.generation, requestId });
    }
  }

  /** Pending calls (tests and leak checks). */
  get pendingCount(): number { return this.pending.size; }

  private settle(requestId: number): boolean {
    const call = this.pending.get(requestId);
    if (!call) return false;
    this.pending.delete(requestId);
    call.detach();
    return true;
  }

  private postQuietly(message: WorkerFontResolveCancel): void {
    try { this.post(message); } catch { /* main is gone; nothing waits there */ }
  }
}

/** An acquire reply outcome, validated; anything malformed is a failure.
 * Each body must be an owned, fixed-length ArrayBuffer within the face
 * ceiling (registration charges it like any `data` face). */
function readAcquireOutcome(value: unknown): FontAcquireOutcome {
  const malformed: FontAcquireOutcome = { status: 'failed', reason: 'malformed font acquire reply' };
  if (value === null || typeof value !== 'object') return malformed;
  const fields = value as { status?: unknown; faces?: unknown; reason?: unknown; violation?: unknown };
  switch (fields.status) {
    case 'faces': {
      if (!Array.isArray(fields.faces) || fields.faces.length > HARD_MAX_FONT_SOURCE_DECLARATIONS) return malformed;
      const faces: ({ url: string; data: ArrayBuffer } | { url: string; reason: string })[] = [];
      for (const entry of fields.faces as unknown[]) {
        const face = (entry ?? {}) as { url?: unknown; data?: unknown; reason?: unknown };
        if (typeof face.url !== 'string') return malformed;
        if (typeof face.reason === 'string') { faces.push({ url: face.url, reason: face.reason.slice(0, MAX_REASON_LENGTH) }); continue; }
        try {
          const checked = validateFontFaceDeclaration({ family: 'acquired', data: face.data });
          faces.push({ url: face.url, data: checked.data as ArrayBuffer });
        } catch (error) {
          if (error instanceof TypeError) return malformed;
          throw error;
        }
      }
      return { status: 'faces', faces };
    }
    case 'failed':
      return typeof fields.reason === 'string' ? { status: 'failed', reason: fields.reason.slice(0, MAX_REASON_LENGTH) } : malformed;
    case 'limit': {
      const violation = fontSourceViolationFromWire(fields.violation);
      return violation ? { status: 'limit', violation } : malformed;
    }
    default:
      return malformed;
  }
}

/** A reply outcome, validated; anything malformed is a failure. */
function readOutcome(value: unknown): FontResolveOutcome {
  const malformed: FontResolveOutcome = { status: 'failed', reason: 'malformed font resolver reply' };
  if (value === null || typeof value !== 'object') return malformed;
  const fields = value as { status?: unknown; families?: unknown; reason?: unknown; violation?: unknown };
  switch (fields.status) {
    case 'families': {
      // The session admits and claims every family's declarations again;
      // only the shape and the counts are bounded here.
      const families = fields.families;
      if (!Array.isArray(families) || families.length > HARD_MAX_FONT_SOURCE_FAMILIES) return malformed;
      const answers: BridgedFamilyAnswer[] = [];
      for (const entry of families as unknown[]) {
        const answer = readFamilyAnswer(entry);
        if (!answer) return malformed;
        answers.push(answer);
      }
      return { status: 'families', families: answers };
    }
    case 'failed':
      return typeof fields.reason === 'string'
        ? { status: 'failed', reason: fields.reason.slice(0, MAX_REASON_LENGTH) }
        : malformed;
    case 'limit': {
      const violation = fontSourceViolationFromWire(fields.violation);
      return violation ? { status: 'limit', violation } : malformed;
    }
    default:
      return malformed;
  }
}

function readFamilyAnswer(value: unknown): BridgedFamilyAnswer | undefined {
  if (value === null || typeof value !== 'object') return undefined;
  const fields = value as { family?: unknown; kind?: unknown; declarations?: unknown; reason?: unknown };
  if (typeof fields.family !== 'string') return undefined;
  const family = fields.family;
  switch (fields.kind) {
    case 'claimed':
      return Array.isArray(fields.declarations) && fields.declarations.length <= HARD_MAX_FONT_SOURCE_DECLARATIONS
        ? { family, kind: 'claimed', declarations: fields.declarations as FontFaceDeclaration[] }
        : undefined;
    case 'failed':
      return typeof fields.reason === 'string'
        ? { family, kind: 'failed', reason: fields.reason.slice(0, MAX_REASON_LENGTH) }
        : undefined;
    case 'absent':
      return { family, kind: 'absent' };
    default:
      return undefined;
  }
}
