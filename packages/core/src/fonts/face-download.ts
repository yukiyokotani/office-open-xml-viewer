/**
 * Library acquisition of `url` faces. Library resource policy, not OOXML
 * behaviour.
 *
 * Face files are normally left to `FontFace.load()` in the painting realm.
 * The library fetches them itself, and hands the bytes to `FontFace` as a
 * `data` face, in two cases only:
 *
 * - a source configured with `credentials: 'omit'` (`fontCssApi`, `fontCss`,
 *   `fontFiles`): CSS font fetching always sends same-origin credentials, so
 *   omitting them needs a request of the library's own, in the painting realm;
 * - a {@link fontResolver} answer of a worker-mode document: the page fetches
 *   the claimed faces' http(s) / blob: URLs and the worker receives the bytes
 *   (`resolver-bridge-host.ts`), so cookies, HTTP authentication and Referer
 *   checks see the page exactly as when the page paints. The viewer-realm
 *   sessions of an XLSX worker-mode workbook acquire through the same cache,
 *   so both realms paint the same bytes from one request.
 *
 * Acquisition runs after a family is claimed (claim precedence is unchanged:
 * a claimed face that cannot be fetched fails like a face that cannot load)
 * and within the painting realm's face-loading ceiling, as the browser's own
 * fetch would. Requests are CORS requests, as CSS font fetching is, at most
 * {@link MAX_CONCURRENT_FACE_DOWNLOADS} at a time per realm.
 *
 * Bytes are charged to the realm's retained-byte ceiling as they arrive, so
 * no batch can outgrow it, and each body is bounded by the face ceiling; a
 * breach rejects with `OoxmlError('font-source-limit')` and aborts the rest of
 * the batch. A network error, a non-OK status or a missed deadline fails that
 * face only. A completed download is shared by every holder of the same
 * (credentials, URL) until the last holder releases it; a failure is not kept.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import { createAbortError } from '../internal/abort.js';
import { HARD_MAX_FONT_SOURCE_FACE_BYTES } from '../worker/resource-policy.generated.js';
import type { FontResolverFailures } from './resolver-failures.js';
import { chargeRetainedBytes, fontSourceId, refundRetainedBytes, SharedTask } from './source-cache.js';
import type { FontFaceDeclaration, FontResolverSource, FontSourceCredentials } from './source.js';

/**
 * Abort reason of a deadline (the face-loading ceiling, a bridged call's
 * timer). Any other abort is a cancellation: the requester is gone, which is
 * no outcome for its document (see {@link acquireResolverFacesInRealm}).
 */
export const FONT_DEADLINE_REASON: unique symbol = Symbol('font deadline');

/** Parallel face downloads per realm (the HTTP/1.1 per-host norm). */
export const MAX_CONCURRENT_FACE_DOWNLOADS = 6;

/** Whether the library acquires `declaration` (an http(s) or blob: URL). */
export function isDownloadableFace(declaration: FontFaceDeclaration): boolean {
  return declaration.url !== undefined && /^(?:https?|blob):/i.test(declaration.url);
}

/** One URL's acquisition outcome. */
export type FaceAcquisition = { readonly data: ArrayBuffer } | { readonly reason: string };

// ---- download slots -------------------------------------------------------------

let active = 0;
const queued: (() => void)[] = [];

/** Wait for a download slot; the returned function frees it. */
async function slot(signal: AbortSignal): Promise<() => void> {
  if (active >= MAX_CONCURRENT_FACE_DOWNLOADS) {
    await new Promise<void>((resolve, reject) => {
      const take = () => { signal.removeEventListener('abort', onAbort); resolve(); };
      const onAbort = () => {
        const at = queued.indexOf(take);
        if (at >= 0) queued.splice(at, 1);
        reject(createAbortError('Font download abandoned'));
      };
      if (signal.aborted) { onAbort(); return; }
      signal.addEventListener('abort', onAbort, { once: true });
      queued.push(take);
    });
  }
  active++;
  let freed = false;
  return () => {
    if (freed) return;
    freed = true;
    active--;
    queued.shift()?.();
  };
}

// ---- one download -----------------------------------------------------------------

interface FaceResponse {
  readonly ok: boolean;
  readonly status?: number;
  readonly headers?: { get(name: string): string | null };
  readonly body?: { getReader?(): ReadableStreamDefaultReader<Uint8Array>; cancel?(): Promise<void> } | null;
  arrayBuffer(): Promise<ArrayBuffer>;
}

/** Read a body within the face ceiling, charging each chunk through
 * `charge` as it arrives. */
async function readBoundedBytes(response: FaceResponse, charge: (bytes: number) => void): Promise<ArrayBuffer> {
  const limit = HARD_MAX_FONT_SOURCE_FACE_BYTES;
  const declared = Number(response.headers?.get('content-length') ?? NaN);
  if (Number.isSafeInteger(declared) && declared > limit) {
    response.body?.cancel?.().catch(() => undefined);
    throw fontSourceLimitError('face', limit, declared);
  }
  const reader = typeof response.body?.getReader === 'function' ? response.body.getReader() : null;
  if (!reader) {
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > limit) throw fontSourceLimitError('face', limit, buffer.byteLength);
    charge(buffer.byteLength);
    return buffer;
  }
  const chunks: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      size += value.byteLength;
      if (size > limit) throw fontSourceLimitError('face', limit, size);
      charge(value.byteLength);
      chunks.push(value);
    }
  } catch (error) {
    reader.cancel().catch(() => undefined);
    throw error;
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
  return bytes.buffer;
}

/**
 * Ownership is by lease. A join takes a lease for its holder when it starts,
 * so an entry another join still waits on is never dropped; a join that
 * fails or whose requester left returns its lease, and a successful one
 * keeps it as the holder's hold. A holder is released once, for good
 * ({@link releaseAcquiredFaces}): its leases are returned and it can never
 * take another, so no continuation that resumes later can hold bytes for it.
 */
const releasedHolders = new WeakSet<object>();

class Entry {
  readonly task: SharedTask<FaceAcquisition>;
  /** Live leases per holder (in-flight joins and kept holds). */
  readonly leases = new Map<object, number>();
  /** Retained bytes charged for this entry (while downloading and kept). */
  charged = 0;
  settled?: FaceAcquisition;

  constructor(produce: (signal: AbortSignal, entry: Entry) => Promise<FaceAcquisition>, onAbandon: (entry: Entry) => void) {
    this.task = new SharedTask((signal) => produce(signal, this), () => onAbandon(this));
  }
}

/** Kept and in-flight acquisitions, by `credentials\0url`. */
const entries = new Map<string, Entry>();

function dropEntry(key: string, entry: Entry): void {
  if (entries.get(key) === entry) entries.delete(key);
  refundRetainedBytes(entry.charged);
  entry.charged = 0;
}

async function download(
  key: string,
  url: string,
  credentials: FontSourceCredentials,
  signal: AbortSignal,
  entry: Entry,
): Promise<FaceAcquisition> {
  const charge = (bytes: number) => {
    chargeRetainedBytes(bytes);
    entry.charged += bytes;
  };
  const free = await slot(signal);
  let result: FaceAcquisition;
  try {
    if (typeof fetch !== 'function') {
      result = { reason: 'fetch is unavailable' };
    } else {
      const response = await fetch(url, { mode: 'cors', credentials, signal }) as FaceResponse;
      if (!response.ok) {
        response.body?.cancel?.().catch(() => undefined);
        result = { reason: `HTTP ${response.status ?? 'error'}` };
      } else {
        const data = await readBoundedBytes(response, charge);
        result = data.byteLength ? { data } : { reason: 'empty font file' };
      }
    }
  } catch (error) {
    dropEntry(key, entry);
    if (error instanceof OoxmlError) throw error;
    if (signal.aborted) throw createAbortError('Font download abandoned');
    return { reason: error instanceof Error ? error.message : String(error) };
  } finally {
    free();
  }
  // Failures are not kept (a later document retries).
  if ('reason' in result || entries.get(key) !== entry) {
    dropEntry(key, entry);
    return result;
  }
  entry.settled = result;
  // Settled after every lease was returned (its last joins left as it
  // finished): nobody keeps it.
  if (entry.leases.size === 0) dropEntry(key, entry);
  return result;
}

function takeLease(entry: Entry, holder: object): void {
  entry.leases.set(holder, (entry.leases.get(holder) ?? 0) + 1);
}

/** Return one lease of `holder`; drops a settled entry nobody leases. */
function returnLease(key: string, entry: Entry, holder: object): void {
  const count = entry.leases.get(holder);
  if (count === undefined) return;
  if (count > 1) entry.leases.set(holder, count - 1);
  else entry.leases.delete(holder);
  if (entry.leases.size === 0 && entry.settled) dropEntry(key, entry);
}

/** One URL for `holder`, joined with any acquisition of it in this realm. */
async function acquireOne(
  url: string,
  credentials: FontSourceCredentials,
  signal: AbortSignal,
  holder: object,
): Promise<FaceAcquisition> {
  // A requester that already left starts nothing and takes no hold.
  if (signal.aborted || releasedHolders.has(holder)) return { reason: 'timed out' };
  const key = `${credentials}\u0000${url}`;
  let entry = entries.get(key);
  if (!entry) {
    // Abandoned before it settled (every waiter left): forget it, so the next
    // request starts afresh; the producer refunds what it charged.
    entry = new Entry(
      (taskSignal, created) => download(key, url, credentials, taskSignal, created),
      (created) => { if (entries.get(key) === created) entries.delete(key); },
    );
    entries.set(key, entry);
  }
  takeLease(entry, holder);
  if (entry.settled) return entry.settled;
  let result: FaceAcquisition;
  try {
    result = await entry.task.join(signal);
  } catch (error) {
    returnLease(key, entry, holder);
    throw error;
  }
  // Released meanwhile: its leases are gone already, and it gets nothing.
  if (releasedHolders.has(holder)) return { reason: 'timed out' };
  // A failure, or a requester that left between the task settling and this
  // continuation: the lease goes back (other leases keep the bytes).
  if (!('data' in result) || signal.aborted) {
    returnLease(key, entry, holder);
    return 'data' in result ? { reason: 'timed out' } : result;
  }
  return result;
}

/** The bytes of a completed acquisition this realm keeps, held for `holder`
 * from now on; undefined when none is kept. */
function keptAcquisition(url: string, credentials: FontSourceCredentials, holder: object): ArrayBuffer | undefined {
  if (releasedHolders.has(holder)) return undefined;
  const entry = entries.get(`${credentials}\u0000${url}`);
  if (!entry?.settled || !('data' in entry.settled)) return undefined;
  takeLease(entry, holder);
  return entry.settled.data;
}

/**
 * Acquire `urls` for `holder` (which keeps every completed download until
 * {@link releaseAcquiredFaces}). A face completed after `deadline` (an
 * absolute `Date.now()` value) fails with `'timed out'`, whatever the timers
 * did. Rejects with `OoxmlError('font-source-limit')` on a ceiling breach,
 * after aborting the rest of the batch; an abort of `signal` fails every
 * unfinished face with `'timed out'`.
 */
export async function acquireFaceUrls(
  urls: Iterable<string>,
  options: {
    readonly credentials: FontSourceCredentials;
    readonly signal: AbortSignal;
    readonly deadline: number;
    readonly holder: object;
  },
): Promise<Map<string, FaceAcquisition>> {
  const { credentials, holder, deadline } = options;
  const batch = new AbortController();
  const onAbort = () => batch.abort();
  options.signal.addEventListener('abort', onAbort, { once: true });
  if (options.signal.aborted) batch.abort();
  try {
    const unique = [...new Set(urls)];
    const results = await Promise.all(unique.map(async (url): Promise<[string, FaceAcquisition]> => {
      try {
        const result = await acquireOne(url, credentials, batch.signal, holder);
        return [url, Date.now() > deadline ? { reason: 'timed out' } : result];
      } catch (error) {
        if (error instanceof OoxmlError) { batch.abort(); throw error; }
        return [url, { reason: 'timed out' }];
      }
    }));
    return new Map(results);
  } finally {
    options.signal.removeEventListener('abort', onAbort);
  }
}

/** Return every acquisition `holder` keeps; idempotent. */
export function releaseAcquiredFaces(holder: object): void {
  releasedHolders.add(holder);
  for (const [key, entry] of [...entries]) {
    if (!entry.leases.delete(holder)) continue;
    if (entry.leases.size === 0 && entry.settled) dropEntry(key, entry);
  }
}

/** Acquisitions this realm keeps or runs (tests, leak checks). */
export function acquiredFaceCount(): number {
  return entries.size;
}

/** `declaration` with its acquired bytes in place of the URL. */
export function withAcquiredBytes(declaration: FontFaceDeclaration, data: ArrayBuffer): FontFaceDeclaration {
  const { family, weight, style, stretch, unicodeRange } = declaration;
  return Object.freeze({
    family,
    data,
    ...(weight === undefined ? {} : { weight }),
    ...(style === undefined ? {} : { style }),
    ...(stretch === undefined ? {} : { stretch }),
    ...(unicodeRange === undefined ? {} : { unicodeRange }),
  });
}

// ---- resolver faces of a worker-mode document --------------------------------

/** One resolver source's claimed URL faces to acquire. */
export interface ResolverFaceRequest {
  readonly source: FontResolverSource;
  /** Position of the source in the document's `fonts`. */
  readonly sourceIndex: number;
  /** Claimed families and the URLs of their faces. */
  readonly families: readonly { readonly family: string; readonly urls: readonly string[] }[];
}

/**
 * Acquires the claimed URL faces of a resolver in the page realm: the worker
 * bridge (`resolver-bridge.ts`) or, for a viewer realm sharing a worker-mode
 * document, {@link acquireResolverFacesInRealm}. Resolves with an outcome per
 * URL; rejects with `OoxmlError('font-source-limit')` on a ceiling breach.
 */
export type ResolverFaceAcquirer = (
  request: ResolverFaceRequest,
  context: ResolverFaceAcquireContext,
) => Promise<ReadonlyMap<string, FaceAcquisition>>;

export interface ResolverFaceAcquireContext {
  /** The face-loading ceiling: aborts with {@link FONT_DEADLINE_REASON} at
   * `deadline`, or as a cancellation when the requester is released. */
  readonly signal: AbortSignal;
  /** Aborts only when the requester is released (never at the deadline): a
   * bridged acquirer waits for the page's verdict until then. */
  readonly cancel: AbortSignal;
  readonly deadline: number;
  readonly holder: object;
}

/** Document-wide record keys of one acquired URL of one resolver source, in
 * the document's resolver failure record (`resolver-failures.ts`): its
 * failure reason, or the marker that it succeeded. */
export function faceAcquisitionFailureKey(sourceId: string, url: string): string {
  return `${sourceId}\u0000acquire\u0000${url}`;
}
function faceAcquisitionSuccessKey(sourceId: string, url: string): string {
  return `${sourceId}\u0000acquired\u0000${url}`;
}

/** The {@link ResolverFaceAcquirer} of a page-realm session that shares a
 * worker-mode document (the XLSX viewer realm): it acquires in this realm,
 * through the document's failure record, exactly as the worker bridge does. */
export function inRealmResolverFaceAcquirer(failures: FontResolverFailures | undefined): ResolverFaceAcquirer {
  return (request, context) => acquireResolverFacesInRealm(failures, fontSourceId(request.source),
    request.families.flatMap(({ urls }) => urls), context);
}

/**
 * Acquire a resolver's claimed URL faces in this (the page) realm, with one
 * outcome per URL for the whole document: the first terminal outcome any
 * requester of the document reaches (the worker bridge, the XLSX viewer
 * realm) is recorded, and every requester adopts it, including one whose own
 * request was still in flight or ended differently (a later timeout adopts
 * the bytes another requester already received and the cache keeps; a
 * success after a recorded failure adopts the failure). A cancellation (the
 * signal aborts before the deadline: the requester is gone) is no outcome
 * and is not recorded.
 */
export async function acquireResolverFacesInRealm(
  failures: FontResolverFailures | undefined,
  sourceId: string,
  urls: readonly string[],
  context: { readonly signal: AbortSignal; readonly deadline: number; readonly holder: object },
): Promise<Map<string, FaceAcquisition>> {
  const credentials = 'same-origin';
  // Classified when the abort happens, by its reason: a deadline is an
  // outcome, anything else (a release, a worker cancel) is not.
  let cancelled = false;
  const onAbort = () => { cancelled = context.signal.reason !== FONT_DEADLINE_REASON; };
  if (context.signal.aborted) onAbort();
  else context.signal.addEventListener('abort', onAbort, { once: true });
  const decided = (url: string): FaceAcquisition | undefined => {
    const reason = failures?.get(faceAcquisitionFailureKey(sourceId, url));
    if (reason !== undefined) return { reason };
    if (!failures?.has(faceAcquisitionSuccessKey(sourceId, url))) return undefined;
    // A requester that is gone (cancelled, released) takes no hold it could
    // never return; it receives no bytes either.
    if (cancelled) return { reason: 'timed out' };
    const data = keptAcquisition(url, credentials, context.holder);
    return data ? { data } : undefined;
  };
  const outcomes = new Map<string, FaceAcquisition>();
  const pending: string[] = [];
  for (const url of new Set(urls)) {
    const outcome = decided(url);
    if (outcome) outcomes.set(url, outcome);
    else pending.push(url);
  }
  let acquired: Map<string, FaceAcquisition>;
  try {
    acquired = await acquireFaceUrls(pending, { credentials, ...context });
  } finally {
    context.signal.removeEventListener('abort', onAbort);
  }
  for (const [url, outcome] of acquired) {
    const adopted = decided(url);
    if (adopted) { outcomes.set(url, adopted); continue; }
    outcomes.set(url, outcome);
    if (!failures || cancelled) continue;
    if ('data' in outcome) failures.set(faceAcquisitionSuccessKey(sourceId, url), '');
    // A success another requester recorded stands (its bytes were released
    // since): this requester's failure is its own, not the document's.
    else if (!failures.has(faceAcquisitionSuccessKey(sourceId, url))) {
      failures.set(faceAcquisitionFailureKey(sourceId, url), outcome.reason);
    }
  }
  return outcomes;
}
