/**
 * Source identity and the per-realm caches font sources share across
 * documents.
 *
 * Library resource policy, not OOXML behaviour. Every cache is keyed by a
 * source's identity ({@link fontSourceId}): a stable hash of its kind and its
 * whole configuration. Two documents configured with equal sources share
 * stylesheet fetches and family resolutions; two sources that differ in
 * anything (origin, template, query, credentials, substitutions, faces) share
 * nothing, so a misbehaving or failing origin can never answer for another
 * (#1497: per-origin isolation, no retry against another origin).
 *
 * Every entry is refcounted by its requesters: a family resolution (present
 * or absent) lives while a session that asked for it is open, and a
 * stylesheet while a family resolution that used it lives
 * ({@link SharedTaskMap.release}). A source's cache itself lives while any
 * session holds the source ({@link acquireSourceCache}). The entry count is
 * therefore bounded by live sessions (at most HARD_MAX_FONT_SOURCE_FAMILIES
 * resolutions each), never by the documents ever opened. A document opened
 * after every earlier requester of a family closed re-issues that request,
 * which the browser HTTP cache answers.
 *
 * Work in the caches is shared through {@link SharedTask}: each waiter joins
 * with its own `AbortSignal`, the underlying fetch is aborted only when every
 * waiter has abandoned it, and a task that fails is evicted so a later session
 * retries the same source (never another).
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { createAbortError } from '../internal/abort.js';
import { HARD_MAX_FONT_SOURCE_RETAINED_BYTES } from '../worker/resource-policy.generated.js';
import { dataUrlDecodedByteLength, type FontFaceDeclaration, type FontSource } from './source.js';
import type { FontSubstitution } from './substitutions.js';

// ---- identity -------------------------------------------------------------------

/** Two-lane 32-bit multiplicative hash (cyrb53 construction), 64 bits of
 * output. The input is integrator configuration, never document content, so
 * collision resistance against an adversary is not required; a collision
 * between two of one application's own sources is negligible at 64 bits. */
class IdentityHash {
  private h1 = 0xdeadbeef;
  private h2 = 0x41c6ce57;

  private unit(value: number): void {
    this.h1 = Math.imul(this.h1 ^ value, 2654435761);
    this.h2 = Math.imul(this.h2 ^ value, 1597334677);
  }

  /** Length-prefixed, so field boundaries cannot be shifted between fields. */
  text(value: string | undefined): this {
    if (value === undefined) { this.unit(0x110000); return this; }
    this.unit(value.length);
    for (let index = 0; index < value.length; index++) this.unit(value.charCodeAt(index));
    return this;
  }

  bytes(value: Uint8Array): this {
    this.unit(0x110001);
    this.unit(value.length);
    for (let index = 0; index < value.length; index++) this.unit(value[index]);
    return this;
  }

  digest(): string {
    let h1 = this.h1, h2 = this.h2;
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (h2 >>> 0).toString(16).padStart(8, '0') + (h1 >>> 0).toString(16).padStart(8, '0');
  }
}

/** Content key of one face's bytes or `data:` URL, as used in registry
 * signatures (a `data:` URL can be tens of MiB; the signature stays short). */
export function fontContentKey(value: ArrayBuffer | string): string {
  if (typeof value === 'string') return `data:${value.length}:${new IdentityHash().text(value).digest()}`;
  return `bytes:${value.byteLength}:${new IdentityHash().bytes(new Uint8Array(value)).digest()}`;
}

function hashSubstitutes(hash: IdentityHash, substitutes: readonly FontSubstitution[] | undefined): void {
  hash.text(substitutes === undefined ? undefined : String(substitutes.length));
  for (const entry of substitutes ?? []) hash.text(entry.for).text(entry.use).text(entry.script);
}

function hashDeclaration(hash: IdentityHash, face: FontFaceDeclaration): void {
  hash.text(face.family).text(face.weight).text(face.style).text(face.stretch).text(face.unicodeRange);
  if (face.data) hash.bytes(new Uint8Array(face.data));
  else hash.text(face.url);
}

const sourceIds = new WeakMap<FontSource, string>();

/**
 * Stable identity of a validated source descriptor in this realm. Descriptors
 * are frozen and their `data` buffers are owned by the library once handed
 * over, so the identity is computed once per descriptor object. A
 * `font-resolver` is identified by its function's id, so documents sharing one
 * resolver function share its resolutions.
 */
export function fontSourceId(source: FontSource): string {
  const cached = sourceIds.get(source);
  if (cached) return cached;
  const hash = new IdentityHash().text(source.protocol).text(source.kind);
  switch (source.kind) {
    case 'font-css-api': {
      hash.text(source.urlTemplate).text(source.query).text(source.credentials ?? 'same-origin');
      const families = source.families ? Object.entries(source.families) : undefined;
      hash.text(families === undefined ? undefined : String(families.length));
      for (const [family, value] of families ?? []) hash.text(family).text(value === true ? undefined : value);
      break;
    }
    case 'font-css':
      hash.text(source.url).text(source.credentials ?? 'same-origin');
      break;
    case 'font-files':
      hash.text(source.credentials ?? 'same-origin').text(String(source.faces.length));
      for (const face of source.faces) hashDeclaration(hash, face);
      break;
    case 'font-resolver':
      hash.text(source.id);
      break;
  }
  hashSubstitutes(hash, source.substitutes);
  const id = `${source.kind}:${hash.digest()}`;
  sourceIds.set(source, id);
  return id;
}

// ---- shared, abandonable work ----------------------------------------------------

/** Abort `controller` when `signal` aborts; returns the detach function. */
function follow(signal: AbortSignal, onAbort: () => void): () => void {
  if (signal.aborted) { onAbort(); return () => undefined; }
  signal.addEventListener('abort', onAbort, { once: true });
  return () => signal.removeEventListener('abort', onAbort);
}

/**
 * One piece of work joined by any number of waiters. The producer starts on
 * the first live join, never before, and receives a signal that aborts as soon
 * as the waiter count returns to zero before the task settled; `onAbandon`
 * then evicts it from its cache so the next waiter starts afresh.
 */
export class SharedTask<T> {
  private readonly controller = new AbortController();
  private promise: Promise<T> | undefined;
  private waiters = 0;
  private settled = false;

  /** Dropped once started: a settled task (possibly cached for the source's
   * lifetime) must keep only its result, never the producer's closure and
   * whatever that captured (a resolver batch, unclaimed font bytes). */
  private produce: ((signal: AbortSignal) => Promise<T>) | undefined;

  constructor(produce: (signal: AbortSignal) => Promise<T>, private readonly onAbandon: () => void) {
    this.produce = produce;
  }

  private start(): Promise<T> {
    const signal = this.controller.signal;
    const produce = this.produce as (signal: AbortSignal) => Promise<T>;
    this.produce = undefined;
    const promise = Promise.resolve()
      .then(() => {
        // Every waiter may have left between the join and this microtask.
        if (signal.aborted) throw createAbortError('Font source request abandoned');
        return produce(signal);
      })
      .finally(() => { this.settled = true; });
    // The task's own rejection is observed by each joiner; never unhandled here.
    promise.catch(() => undefined);
    return promise;
  }

  /** Await the task, or reject with an `AbortError` as soon as `signal` aborts.
   * Joining with an aborted signal neither starts nor keeps the task. */
  join(signal: AbortSignal): Promise<T> {
    if (signal.aborted || this.controller.signal.aborted) {
      if (this.waiters === 0 && !this.settled) this.abandon();
      return Promise.reject(createAbortError('Font source request abandoned'));
    }
    this.waiters++;
    const promise = this.promise ??= this.start();
    return new Promise<T>((resolve, reject) => {
      let done = false;
      const leave = () => {
        if (done) return;
        done = true;
        this.waiters--;
      };
      const detach = follow(signal, () => {
        if (done) return;
        leave();
        if (this.waiters === 0 && !this.settled) this.abandon();
        reject(createAbortError('Font source request abandoned'));
      });
      promise.then(
        (value) => { detach(); if (!done) { leave(); resolve(value); } },
        (error: unknown) => { detach(); if (!done) { leave(); reject(error); } },
      );
    });
  }

  private abandon(): void {
    this.produce = undefined;
    this.controller.abort();
    this.onAbandon();
  }
}

// ---- realm byte accounting -------------------------------------------------------

/** Decoded bytes of font data the library keeps alive in this realm: cached
 * resolver buffers and cached `data:` URLs (charged by their cache entries)
 * plus registered `data` / `data:` faces (charged by their registrations, see
 * `source-session.ts`). Both are real copies in memory, so both count against
 * `HARD_MAX_FONT_SOURCE_RETAINED_BYTES`. */
let retainedTotal = 0;

/** Charge `bytes` or reject with `OoxmlError('font-source-limit')`. */
export function chargeRetainedBytes(bytes: number): void {
  if (bytes <= 0) return;
  if (retainedTotal + bytes > HARD_MAX_FONT_SOURCE_RETAINED_BYTES) {
    throw fontSourceLimitError('retained', HARD_MAX_FONT_SOURCE_RETAINED_BYTES, retainedTotal + bytes);
  }
  retainedTotal += bytes;
}

export function refundRetainedBytes(bytes: number): void {
  if (bytes > 0) retainedTotal -= bytes;
}

/** Bytes this realm currently retains for font sources (tests, leak checks). */
export function retainedFontSourceBytes(): number {
  return retainedTotal;
}

/** Decoded bytes a declaration list holds: `data` buffers (each distinct
 * buffer once) and `data:` URLs. */
export function declarationBytes(declarations: readonly FontFaceDeclaration[]): number {
  const buffers = new Set<ArrayBuffer>();
  let bytes = 0;
  for (const declaration of declarations) {
    if (declaration.data) {
      if (!buffers.has(declaration.data)) { buffers.add(declaration.data); bytes += declaration.data.byteLength; }
    } else if (declaration.url && /^\s*data:/i.test(declaration.url)) {
      bytes += dataUrlDecodedByteLength(declaration.url) ?? 0;
    }
  }
  return bytes;
}

/**
 * A keyed map of shared tasks. Rejections and results `evict` accepts are
 * removed once settled, so failures are never cached. A kept result is charged
 * `weigh(result)` retained bytes for as long as the map keeps it; a charge
 * beyond the realm ceiling rejects every joiner and keeps nothing.
 */
export class SharedTaskMap<T> {
  private readonly tasks = new Map<string, SharedTask<T>>();
  /** Bytes charged per task (not per key): a retry after an eviction is a
   * different task, so charge and refund always pair one-to-one. */
  private readonly charged = new Map<SharedTask<T>, number>();
  /** Results the map keeps, by task (see {@link SharedTaskMap.holdKept}). */
  private readonly values = new Map<SharedTask<T>, T>();
  /** Who asked for each key (sessions for family resolutions, family
   * resolutions for stylesheets). An entry is kept only while at least one
   * holder needs it, so the map is bounded by its live holders' demand. */
  private readonly holders = new Map<string, Set<object>>();
  private readonly held = new Map<object, Set<string>>();

  constructor(
    private readonly evict: (value: T) => boolean = () => false,
    private readonly weigh: (value: T) => number = () => 0,
    /** Called once a task leaves the map (and again if its producer finishes
     * afterwards), so work it held elsewhere is released too. */
    private readonly onDrop: (task: object) => void = () => undefined,
  ) {}

  /** Forget `task` (if it is still the entry for `key`) and refund its
   * charge. Every removal path goes through here. */
  private drop(key: string, task: SharedTask<T>): void {
    if (this.tasks.get(key) === task) this.tasks.delete(key);
    this.values.delete(task);
    const bytes = this.charged.get(task);
    if (bytes !== undefined) {
      this.charged.delete(task);
      refundRetainedBytes(bytes);
    }
    this.onDrop(task);
  }

  private hold(key: string, holder: object): void {
    let keyHolders = this.holders.get(key);
    if (!keyHolders) { keyHolders = new Set(); this.holders.set(key, keyHolders); }
    keyHolders.add(holder);
    let keys = this.held.get(holder);
    if (!keys) { keys = new Set(); this.held.set(holder, keys); }
    keys.add(key);
  }

  /**
   * Join (or start) the task for `key` on behalf of `holder`. The producer
   * receives the task's signal and the task itself, which it uses as the
   * holder of whatever it joins in other maps.
   */
  join(
    key: string,
    produce: (signal: AbortSignal, task: object) => Promise<T>,
    signal: AbortSignal,
    holder: object,
  ): Promise<T> {
    // Never create or hold work for a waiter that has already left.
    if (signal.aborted) return Promise.reject(createAbortError('Font source request abandoned'));
    this.hold(key, holder);
    let task = this.tasks.get(key);
    if (!task) {
      const created: SharedTask<T> = new SharedTask(async (taskSignal) => {
        try {
          const value = await produce(taskSignal, created);
          // Dropped while producing: nothing keeps the value, charge nothing,
          // and release what the producer held after the drop.
          if (this.tasks.get(key) !== created) { this.onDrop(created); return value; }
          if (this.evict(value)) { this.drop(key, created); return value; }
          const bytes = this.weigh(value);
          chargeRetainedBytes(bytes);
          if (bytes > 0) this.charged.set(created, bytes);
          this.values.set(created, value);
          return value;
        } catch (error) {
          this.drop(key, created);
          throw error;
        }
      }, () => this.drop(key, created));
      task = created;
      this.tasks.set(key, task);
    }
    return task.join(signal);
  }

  /** `holder` no longer needs anything here: every entry it alone held is
   * dropped (and refunded). */
  release(holder: object): void {
    const keys = this.held.get(holder);
    if (!keys) return;
    this.held.delete(holder);
    for (const key of keys) {
      const keyHolders = this.holders.get(key);
      keyHolders?.delete(holder);
      if (keyHolders && keyHolders.size > 0) continue;
      this.holders.delete(key);
      const task = this.tasks.get(key);
      if (task) this.drop(key, task);
    }
  }

  /**
   * The settled result kept for `key`, now also held by `holder` (as a join
   * would hold it), or undefined when none is kept (in flight, absent). The
   * only way to read a kept result: an answer taken from the cache always
   * keeps its entry alive for whoever received it.
   */
  holdKept(key: string, holder: object): T | undefined {
    const task = this.tasks.get(key);
    const value = task ? this.values.get(task) : undefined;
    if (value !== undefined) this.hold(key, holder);
    return value;
  }

  /** Whether a task (in flight or settled and kept) exists for `key`. */
  has(key: string): boolean { return this.tasks.has(key); }

  /** Entries currently kept (tests and leak checks). */
  get size(): number { return this.tasks.size; }

  clear(): void {
    for (const [key, task] of [...this.tasks]) this.drop(key, task);
    // Abandoned-but-charged tasks are refunded by drop() already; nothing
    // charged can remain without an entry.
    for (const bytes of this.charged.values()) refundRetainedBytes(bytes);
    this.charged.clear();
    this.values.clear();
    this.holders.clear();
    this.held.clear();
  }
}

// ---- per-source cache ------------------------------------------------------------

/** Parsed stylesheet: its declarations and the final URL after redirects. */
export interface FontStylesheet {
  readonly ok: true;
  readonly declarations: readonly FontFaceDeclaration[];
  readonly finalUrl: string;
}

/** A transport or content failure: the families keep their host fallback. */
export interface FontSourceFailure {
  readonly ok: false;
  readonly reason: string;
}

/** How one source answered one demanded family. */
export type FamilyResolution =
  /** The source declares faces for the family; the family is now final. */
  | {
    readonly kind: 'claimed';
    /** CSS family the declarations register under. */
    readonly served: string;
    /** Present when a declared substitute serves the demanded family. */
    readonly substitution?: FontSubstitution;
    readonly declarations: readonly FontFaceDeclaration[];
  }
  /** The source attempted the family and failed; the family is also final
   * (host fallback, no other source is consulted). */
  | { readonly kind: 'failed'; readonly reason: string }
  /** The source does not serve the family; the next source is consulted. */
  | { readonly kind: 'absent' };

/** Caches one source identity shares across the sessions that hold it. */
export interface SourceCache {
  readonly id: string;
  /** Stylesheets keyed by request URL. */
  readonly stylesheets: SharedTaskMap<FontStylesheet | FontSourceFailure>;
  /** Resolutions keyed by lower-cased demanded family. */
  readonly families: SharedTaskMap<FamilyResolution>;
  /** Family index of a `font-files` source (computed once; it references
   * only the descriptor's own declarations). */
  index?: ReadonlyMap<string, readonly FontFaceDeclaration[]>;
  release(): void;
}

interface CacheEntry {
  cache: SourceCache;
  refs: number;
}

const caches = new Map<string, CacheEntry>();

/** Hold the shared cache of one source identity until `release()`. */
export function acquireSourceCache(source: FontSource): SourceCache {
  const id = fontSourceId(source);
  const existing = caches.get(id);
  if (existing) {
    existing.refs++;
    return handle(existing);
  }
  // Cached content is charged where the cache is its only holder: inline
  // data: faces of a stylesheet, and the buffers a resolver returned. (A
  // font-files index references the descriptor's own buffers.) Entry counts
  // are bounded by holders: a family resolution lives while a session that
  // requested it is open (at most HARD_MAX_FONT_SOURCE_FAMILIES per session),
  // a stylesheet while a family resolution that used it lives.
  const stylesheets = new SharedTaskMap<FontStylesheet | FontSourceFailure>((value) => !value.ok,
    (value) => (value.ok ? declarationBytes(value.declarations) : 0));
  const families = new SharedTaskMap<FamilyResolution>((value) => value.kind === 'failed',
    (value) => (source.kind === 'font-resolver' && value.kind === 'claimed'
      ? declarationBytes(value.declarations) : 0),
    (task) => stylesheets.release(task));
  const entry: CacheEntry = {
    refs: 1,
    cache: { id, stylesheets, families, release: () => undefined },
  };
  caches.set(id, entry);
  return handle(entry);
}

/** A per-holder view whose `release()` is idempotent. */
function handle(entry: CacheEntry): SourceCache {
  let released = false;
  const { cache } = entry;
  const view: SourceCache = {
    id: cache.id,
    stylesheets: cache.stylesheets,
    families: cache.families,
    get index() { return cache.index; },
    set index(value) { cache.index = value; },
    release() {
      if (released) return;
      released = true;
      // Family resolutions only this holder requested go now, with the
      // stylesheets they alone used.
      cache.families.release(view);
      entry.refs--;
      if (entry.refs > 0 || caches.get(cache.id) !== entry) return;
      caches.delete(cache.id);
      // Drop every structure that holds font content together with the
      // refund: the task maps (charged) and the family index (it references
      // the cached stylesheet's declarations, inline data: faces included).
      cache.stylesheets.clear();
      cache.families.clear();
      cache.index = undefined;
    },
  };
  return view;
}

/** Number of source identities with live holders (tests and leak checks). */
export function liveSourceCacheCount(): number {
  return caches.size;
}

/** Cached stylesheet and family entries across all live sources (tests and
 * leak checks). */
export function sourceCacheEntryCount(): number {
  let count = 0;
  for (const { cache } of caches.values()) count += cache.stylesheets.size + cache.families.size;
  return count;
}
