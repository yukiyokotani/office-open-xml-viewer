/**
 * `FontSourceSession`: one document's font sources in one painting realm.
 *
 * One generic loader for every source kind (descriptor-only sources and
 * application resolvers alike). Library resource policy, not OOXML behaviour: ECMA-376 names
 * faces but leaves their provision to the host. Each {@link
 * FontSourceSession.ensure} takes a demand delta and
 *
 * 1. **resolves** every newly demanded family (authored and library fallback)
 *    by walking the sources in option order. The first source that declares a
 *    face for the family, directly or through one of its declared
 *    substitutes, claims it. A source that attempts a family and fails (a
 *    stylesheet that cannot be fetched or yields no rule, a resolver that
 *    throws, the 15 s ceiling) also makes the family final: it keeps its host
 *    fallback and no later source is consulted, so a failing mirror never
 *    falls through to another origin (#1497 R5) and first paint is
 *    deterministic. Only a source that does not serve the family (an
 *    allow-list miss, a family absent from a stylesheet or file list) passes
 *    it on;
 *    The claimed `url` faces of a source with `credentials: 'omit'`, and of
 *    a resolver when the session acquires resolver faces (the worker bridge,
 *    an XLSX viewer realm of a worker-mode workbook), are then fetched by the
 *    library (`face-download.ts`) within the face-loading ceiling; a face
 *    that cannot be fetched fails like a face that cannot load;
 * 2. **admits** the claimed declarations against the generated ceilings
 *    (families per session, declarations per session, retained bytes per
 *    realm). A breach rejects with `OoxmlError('font-source-limit')`, releases
 *    the session and is rethrown by every later ensure: it is a misconfigured
 *    or hostile source, not network weather;
 * 3. **registers** every declaration of a claimed family, in declaration
 *    order, through the refcounted per-realm registry (`font-registry.ts`).
 *    Registering excluded subsets too keeps overlapping `unicode-range`
 *    precedence independent of the order demand arrives in;
 * 4. **loads** (`FontFace.load()`) only faces whose `unicode-range`
 *    intersects the demanded scalars (every face for `'all'`, and faces
 *    without a range always). Canvas never triggers a load itself, so an
 *    unloaded face paints nothing; later deltas load further subsets;
 * 5. **snapshots** {@link LoadedFontSources}: the faces that actually loaded,
 *    the substitutions whose served face loaded, and regular-name aliases.
 *
 * Failures are reported with one `console.warn` per newly failed family set;
 * nothing else is surfaced (no strict mode, no feedback API). Ensures are
 * serialised; stylesheets and resolutions are shared across sessions of the
 * same source identity (`source-cache.ts`), and faces across sessions of the
 * same realm (`font-registry.ts`). {@link FontSourceSession.release} returns
 * every reference the session took, exactly once.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import { createAbortError } from '../internal/abort.js';
import {
  HARD_MAX_FONT_DEMAND_SCALARS,
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { parseUnicodeRange, resolveCssSource } from './css-source.js';
import {
  FONT_DEADLINE_REASON,
  acquireFaceUrls,
  isDownloadableFace,
  releaseAcquiredFaces,
  withAcquiredBytes,
  type FaceAcquisition,
  type ResolverFaceAcquirer,
  type ResolverFaceRequest,
} from './face-download.js';
import { resolveFontFiles } from './files-source.js';
import { concludeFamilies } from './resolver-failures.js';
import {
  claimResolverCandidates,
  resolveFamiliesThroughCache,
  runResolverInRealm,
  substituteScriptsOf,
  type FontResolverFailures,
  type FontResolverRunner,
} from './resolver-resolution.js';
import { activeFontSet, FONT_LOAD_CEILING_MS, releaseFaces, retainFace } from './font-registry.js';
import {
  acquireSourceCache,
  chargeRetainedBytes,
  fontContentKey,
  refundRetainedBytes,
  retainedFontSourceBytes,
  type FamilyResolution,
  type SourceCache,
} from './source-cache.js';
import {
  assertFontSourceDescriptors,
  dataUrlDecodedByteLength,
  type FontFaceDeclaration,
  type FontResolverSource,
  type FontSource,
} from './source.js';
import type { FontSubstituteScript } from './substitute-script.js';
import {
  fontFamilyKey,
  loadedRegularAliases,
  type FontSubstituteScripts,
  type FontSubstitution,
} from './substitutions.js';

/** One face a source supplied that loaded in this realm. */
export interface LoadedSourceFace {
  /** CSS family the face is registered under. */
  readonly family: string;
  readonly weight: string;
  readonly style: string;
  readonly stretch: string;
  readonly unicodeRange: string;
  /** Index of the supplying source in the session's `sources`. */
  readonly sourceIndex: number;
  readonly face: FontFace;
}

/** What the renderers consume after each ensure. */
export interface LoadedFontSources {
  /** Loaded faces, in registration order. */
  readonly faces: readonly LoadedSourceFace[];
  /** Lower-cased demanded family → served family, only for substitutions at
   * least one of whose served faces loaded. */
  readonly substitutions: ReadonlyMap<string, { readonly family: string; readonly script?: FontSubstituteScript }>;
  /** `loadedRegularAliases` over the loaded faces ('lato regular' → 'Lato'). */
  readonly regularAliases: ReadonlyMap<string, string>;
}

export type { FontResolverFailures, FontResolverRunner, ResolverFaceAcquirer, ResolverFaceRequest };

export interface FontSourceSessionOptions {
  /** The `fonts` option value; validated here. */
  readonly sources: readonly FontSource[];
  /** Target set; defaults to this realm's (`document.fonts` / `self.fonts`). */
  readonly fontSet?: FontFaceSet | null;
  /** Aborting releases the session and rejects pending and later ensures. */
  readonly signal?: AbortSignal;
  readonly runResolver?: FontResolverRunner;
  /** Resolver failures shared by every requester of one document (XLSX: the
   * viewer-realm sessions and the worker bridge); see `resolver-resolution.ts`. */
  readonly resolverFailures?: FontResolverFailures;
  /** The runner enforces the font ceiling itself and always settles (the
   * worker bridge, whose main side times calls out and records the outcome
   * for the whole document), so resolver waits ignore this realm's ceiling. */
  readonly resolverCeilingOwnedByRunner?: boolean;
  /** Fetch the claimed http(s) / blob: faces of resolver sources through
   * this acquirer instead of `FontFace.load()` (see {@link ResolverFaceAcquirer}). */
  readonly acquireResolverFaces?: ResolverFaceAcquirer;
  /** The acquirer hands over fresh buffers this session alone keeps (the
   * worker bridge's transferred copies), so the session charges each one
   * against the realm ceiling while it keeps it. Otherwise the realm's
   * acquisition cache owns and charges them. */
  readonly acquiredFacesOwnedBySession?: boolean;
}


const EMPTY_SNAPSHOT: LoadedFontSources = Object.freeze({
  faces: Object.freeze([]),
  substitutions: new Map(),
  regularAliases: new Map(),
});

// ---- demand normalisation -----------------------------------------------------------

/**
 * Library-internal demand of one document (or one progressive delta) on its
 * sessions. A resolver receives only the families (`FontDemand`); the rest
 * stays with the library.
 */
export interface FontSessionDemand {
  /** Authored family names not proven installed on the host (trimmed, original case). */
  readonly families: readonly string[];
  /** Library fallback families the renderer puts in CSS tails for the scripts
   * present in the text. */
  readonly fallbackFamilies: readonly string[];
  /** Unicode scalars seen in text, sorted, or 'all'. Drives unicode-range laziness. */
  readonly scalars: 'all' | readonly number[];
  /**
   * Substitute scripts (core `substitute-script.ts`) each family's painted
   * text proves, keyed by lower-cased family. When present, a script-scoped
   * substitution is consulted for a family only if its script is listed for
   * that family (a family absent from the record proves none); the family's
   * own faces are always consulted. Omitted: no restriction.
   */
  readonly substituteScripts?: Readonly<Record<string, readonly FontSubstituteScript[]>>;
}

interface NormalizedDemand {
  /** Authored then fallback families, trimmed, first spelling per key, with
   * the substitute scripts their text proves. */
  readonly names: readonly DemandedName[];
  /** Sorted, unique, or 'all'. */
  readonly scalars: 'all' | readonly number[];
}

interface DemandedName {
  readonly name: string;
  readonly fallback: boolean;
  readonly scripts: FontSubstituteScripts;
}

function familyList(value: unknown, location: string): string[] {
  if (value === undefined) return [];
  if (!Array.isArray(value)) throw new TypeError(`${location} must be an array of family names`);
  return value.map((name: unknown, index) => {
    if (typeof name !== 'string') throw new TypeError(`${location}[${index}] must be a string`);
    return name.trim();
  });
}

/** Validate a demand delta and normalise its scalars (sorted, unique; more
 * than `HARD_MAX_FONT_DEMAND_SCALARS` becomes `'all'`, the conservative
 * answer). Library-produced demand never fails this; it guards the contract. */
export function normalizeFontDemand(demand: FontSessionDemand): NormalizedDemand {
  if (!demand || typeof demand !== 'object') throw new TypeError('font demand must be an object');
  const names: DemandedName[] = [];
  const seen = new Set<string>();
  const scriptsOf = substituteScriptsOf(demand.substituteScripts);
  const add = (list: string[], fallback: boolean) => {
    for (const name of list) {
      const key = fontFamilyKey(name);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      names.push({ name, fallback, scripts: scriptsOf(name) });
    }
  };
  add(familyList(demand.families, 'font demand families'), false);
  add(familyList(demand.fallbackFamilies, 'font demand fallbackFamilies'), true);
  let scalars: 'all' | readonly number[];
  if (demand.scalars === 'all') {
    scalars = 'all';
  } else {
    if (!Array.isArray(demand.scalars)) throw new TypeError("font demand scalars must be 'all' or an array");
    const unique = new Set<number>();
    for (const scalar of demand.scalars) {
      if (!Number.isInteger(scalar) || scalar < 0 || scalar > 0x10ffff) {
        throw new TypeError('font demand scalars must be integers in U+0000..U+10FFFF');
      }
      unique.add(scalar);
    }
    scalars = unique.size > HARD_MAX_FONT_DEMAND_SCALARS ? 'all' : [...unique].sort((a, b) => a - b);
  }
  return { names, scalars };
}

// ---- helpers ---------------------------------------------------------------------------

function rangeIntersects(ranges: readonly (readonly [number, number])[], sorted: readonly number[]): boolean {
  for (const [lo, hi] of ranges) {
    let left = 0, right = sorted.length;
    while (left < right) {
      const mid = (left + right) >>> 1;
      if (sorted[mid] < lo) left = mid + 1;
      else right = mid;
    }
    if (left < sorted.length && sorted[left] <= hi) return true;
  }
  return false;
}

/** A CSS `url()` for a validated absolute URL (data: URLs are kept verbatim
 * by validation and may contain characters a CSS string must escape). */
function cssUrl(url: string): string {
  return `url("${url.replace(/[\\"\n\r\f]/g, (char) =>
    char === '\\' || char === '"' ? `\\${char}` : `\\${char.charCodeAt(0).toString(16)} `)}")`;
}

const isAbortError = (error: unknown) => error instanceof Error && error.name === 'AbortError';

/** A signal aborted by `parent` or after `ms`; `dispose()` clears the timer. */
function ceilingSignal(parent: AbortSignal, ms: number): {
  signal: AbortSignal; expired: () => boolean; remaining: () => number; dispose: () => void;
} {
  const controller = new AbortController();
  let expired = false;
  const deadline = Date.now() + ms;
  const onAbort = () => controller.abort();
  parent.addEventListener('abort', onAbort, { once: true });
  const timer = setTimeout(() => { expired = true; controller.abort(FONT_DEADLINE_REASON); }, ms);
  if (parent.aborted) controller.abort();
  return {
    signal: controller.signal,
    expired: () => expired,
    remaining: () => (expired ? 0 : Math.max(0, deadline - Date.now())),
    // Abandon whatever still waits on this phase (a sibling rejected first).
    dispose: () => { clearTimeout(timer); parent.removeEventListener('abort', onAbort); controller.abort(); },
  };
}

/** Await `promise`, or stop waiting when `signal` aborts (the work itself
 * continues; its result is then ignored). */
function until<T>(promise: Promise<T>, signal: AbortSignal): Promise<T | undefined> {
  if (signal.aborted) return Promise.resolve(undefined);
  return new Promise((resolve, reject) => {
    const onAbort = () => resolve(undefined);
    signal.addEventListener('abort', onAbort, { once: true });
    promise.then(
      (value) => { signal.removeEventListener('abort', onAbort); resolve(value); },
      (error: unknown) => { signal.removeEventListener('abort', onAbort); reject(error); },
    );
  });
}

// ---- registration byte accounting -----------------------------------------------

/** Decoded bytes of `data` faces and `data:` URL faces registered in this
 * realm, per FontFaceSet and signature (a face shared by two documents is
 * charged once; each realm's set holds its own copy). The realm total is
 * shared with the source caches (`source-cache.ts`). */
const registeredBytes = new WeakMap<FontFaceSet, Map<string, { bytes: number; refs: number }>>();

function chargeBytes(set: FontFaceSet, signature: string, bytes: number): void {
  if (bytes === 0) return;
  let bySignature = registeredBytes.get(set);
  const existing = bySignature?.get(signature);
  if (existing) { existing.refs++; return; }
  chargeRetainedBytes(bytes);
  if (!bySignature) { bySignature = new Map(); registeredBytes.set(set, bySignature); }
  bySignature.set(signature, { bytes, refs: 1 });
}

function refundBytes(set: FontFaceSet, signature: string): void {
  const bySignature = registeredBytes.get(set);
  const entry = bySignature?.get(signature);
  if (!bySignature || !entry) return;
  if (--entry.refs > 0) return;
  bySignature.delete(signature);
  refundRetainedBytes(entry.bytes);
}

export { retainedFontSourceBytes };

// ---- session ------------------------------------------------------------------------

interface OwnedFace {
  readonly signature: string;
  readonly face: FontFace;
  readonly declaration: FontFaceDeclaration;
  readonly sourceIndex: number;
  /** null: no (or an unparseable) unicode-range, so the face covers all. */
  readonly ranges: readonly (readonly [number, number])[] | null;
  readonly charged: boolean;
  /** Demanded families this face serves (for warnings). */
  readonly families: Set<string>;
  readiness?: Promise<void>;
  /** Settled load outcome; undefined while loading. */
  outcome?: boolean;
}

type FamilyState =
  | { readonly kind: 'claimed'; readonly name: string; readonly sourceIndex: number; readonly served: string;
    readonly substitution?: FontSubstitution; readonly declarations: readonly FontFaceDeclaration[];
    readonly faces: OwnedFace[];
    /** A library-fetched face of the family failed (the family is warned). */
    readonly downloadFailure?: string }
  | { readonly kind: 'failed'; readonly name: string; readonly reason: string }
  /** No source declared a face for the family under substitution scope
   * `scripts`. A later demand that proves a wider scope resolves it again. */
  | { readonly kind: 'absent'; readonly name: string; readonly scripts: FontSubstituteScripts };

/** The union of two substitution scopes. */
function widenScope(a: FontSubstituteScripts, b: FontSubstituteScripts): FontSubstituteScripts {
  if (a === 'all' || b === 'all') return 'all';
  return Object.freeze([...new Set([...a, ...b])].sort());
}

/** Whether scope `b` admits a script-scoped substitution `a` does not. */
function scopeWidens(a: FontSubstituteScripts, b: FontSubstituteScripts): boolean {
  if (a === 'all') return false;
  return b === 'all' || b.some((script) => !a.includes(script));
}

export class FontSourceSession {
  // Cleared on release: descriptors hold font bytes (fontFiles data,
  // inline data: URLs) and the resolver runner may capture application data.
  private sources: readonly FontSource[];
  private caches: readonly SourceCache[];
  private fontSet: FontFaceSet | null;
  private runResolver: FontResolverRunner;
  private resolverFailures: FontResolverFailures | undefined;
  private acquireResolverFaces: ResolverFaceAcquirer | undefined;
  private readonly ownsAcquiredFaces: boolean;
  /** Resolver faces this session acquired, by `sourceIndex\0url`: a later
   * delta claiming the same face reuses the bytes instead of asking again. */
  private readonly acquiredFaces = new Map<string, ArrayBuffer>();
  private acquiredFaceBytes = 0;
  private readonly runnerOwnsCeiling: boolean;
  private readonly controller = new AbortController();
  private detachSignal: () => void = () => undefined;
  private closed = false;
  private abortedByCaller = false;
  private failure: unknown = undefined;
  private queue: Promise<unknown> = Promise.resolve();
  private readonly families = new Map<string, FamilyState>();
  private readonly owned = new Map<string, OwnedFace>();
  private all = false;
  private readonly points = new Set<number>();
  private readonly warned = new Set<string>();
  private snapshot: LoadedFontSources = EMPTY_SNAPSHOT;

  constructor(options: FontSourceSessionOptions) {
    this.sources = assertFontSourceDescriptors(options.sources);
    this.fontSet = options.fontSet === undefined ? activeFontSet() : options.fontSet;
    this.runResolver = options.runResolver ?? runResolverInRealm;
    this.resolverFailures = options.resolverFailures;
    this.runnerOwnsCeiling = options.resolverCeilingOwnedByRunner === true;
    this.acquireResolverFaces = options.acquireResolverFaces;
    this.ownsAcquiredFaces = options.acquiredFacesOwnedBySession === true;
    this.caches = this.sources.map(acquireSourceCache);
    const signal = options.signal;
    if (signal) {
      const onAbort = () => { this.abortedByCaller = true; this.release(); };
      if (signal.aborted) onAbort();
      else {
        signal.addEventListener('abort', onAbort, { once: true });
        this.detachSignal = () => signal.removeEventListener('abort', onAbort);
      }
    }
  }

  /** The snapshot after the latest completed ensure. */
  get loaded(): LoadedFontSources { return this.snapshot; }

  /**
   * Resolve, register and load one demand delta (families already handled by
   * this session are not resolved again; new scalars may load further subsets
   * of claimed families). Resolves with the snapshot; rejects with
   * `OoxmlError('font-source-limit')` on a ceiling breach and with an
   * `AbortError` once the caller's signal aborted. After {@link release} it
   * resolves with an empty snapshot.
   */
  ensure(demand: FontSessionDemand): Promise<LoadedFontSources> {
    let normalized: NormalizedDemand;
    try { normalized = normalizeFontDemand(demand); }
    catch (error) { return Promise.reject(error); }
    const task = this.queue.then(() => this.ensureNow(normalized));
    // The queue only orders ensures: it must never hold a settled snapshot
    // (its FontFaces) past the caller's own reference.
    this.queue = task.then(() => undefined, () => undefined);
    return task;
  }

  /** Release every face, byte charge and cache reference; idempotent. */
  release(): void {
    if (this.closed) return;
    this.closed = true;
    this.controller.abort();
    this.detachSignal();
    releaseFaces([...this.owned.values()].map((entry) => entry.face));
    if (this.fontSet) for (const entry of this.owned.values()) if (entry.charged) refundBytes(this.fontSet, entry.signature);
    this.owned.clear();
    this.families.clear();
    this.points.clear();
    for (const cache of this.caches) cache.release();
    releaseAcquiredFaces(this);
    this.acquiredFaces.clear();
    refundRetainedBytes(this.acquiredFaceBytes);
    this.acquiredFaceBytes = 0;
    // A released session object the application still references must not
    // keep the caches (and what they index) reachable.
    this.caches = [];
    // Nothing that holds bytes or FontFaces outlives the release: the source
    // descriptors, the target set, the resolver runner, the warned names.
    this.sources = [];
    this.fontSet = null;
    this.runResolver = releasedRunner;
    this.resolverFailures = undefined;
    this.acquireResolverFaces = undefined;
    this.warned.clear();
    this.snapshot = EMPTY_SNAPSHOT;
  }

  private settledState(): LoadedFontSources {
    if (this.abortedByCaller) throw createAbortError('Font source session aborted');
    if (this.failure !== undefined) throw this.failure;
    return this.snapshot;
  }

  private async ensureNow(demand: NormalizedDemand): Promise<LoadedFontSources> {
    if (this.closed || this.failure !== undefined) return this.settledState();
    if (!this.fontSet || typeof FontFace === 'undefined' || this.sources.length === 0) return this.snapshot;
    try {
      // A family is resolved once, except that an absent family is resolved
      // again when a later delta proves a substitution scope its earlier
      // resolution excluded (progressive loading: Latin text on one slide,
      // Arabic text in the same family on a later one), under the union of
      // both scopes. A claimed or failed family is final.
      const fresh: DemandedName[] = [];
      let added = 0;
      for (const entry of demand.names) {
        const prior = this.families.get(fontFamilyKey(entry.name));
        if (!prior) { fresh.push(entry); added++; continue; }
        if (prior.kind === 'absent' && scopeWidens(prior.scripts, entry.scripts)) {
          fresh.push({ ...entry, scripts: widenScope(prior.scripts, entry.scripts) });
        }
      }
      if (this.families.size + added > HARD_MAX_FONT_SOURCE_FAMILIES) {
        throw fontSourceLimitError('families', HARD_MAX_FONT_SOURCE_FAMILIES, this.families.size + added);
      }
      const delta = this.addScalars(demand.scalars);
      const resolved = fresh.length ? await this.resolve(fresh) : new Map<string, FamilyState>();
      if (this.closed) return this.settledState();
      // One face-loading ceiling covers acquisition and loading, as the
      // browser's own fetch inside FontFace.load() would.
      const ceiling = ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS);
      try {
        return await this.acquireRegisterLoad(fresh, resolved, delta, ceiling);
      } finally {
        ceiling.dispose();
      }
    } catch (error) {
      if (this.closed) return this.settledState();
      this.failure = error;
      this.release();
      throw error;
    }
  }

  private async acquireRegisterLoad(
    fresh: readonly DemandedName[],
    resolved: Map<string, FamilyState>,
    delta: 'all' | readonly number[],
    ceiling: ReturnType<typeof ceilingSignal>,
  ): Promise<LoadedFontSources> {
    const acquired = await this.acquire(resolved, ceiling);
    if (this.closed) return this.settledState();
    // A bridged acquisition answers with the page's verdict, which may arrive
    // after this realm's ceiling (main owns that deadline): the faces it
    // delivered still get a bounded wait for decoding, so this realm paints
    // what the page decided, as another realm of the document does.
    // By the absolute deadline, not the ceiling's timer, which may not have
    // run yet when a busy realm receives the reply.
    const loadCeiling = acquired && ceiling.remaining() === 0
      ? ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS) : undefined;
    const registered = new Set<OwnedFace>();
    const failures = new Map<string, string>();
    for (const { name, scripts } of fresh) {
      const key = fontFamilyKey(name);
      const state = resolved.get(key) ?? { kind: 'absent', name, scripts };
      this.families.set(key, state);
      if (state.kind === 'failed') failures.set(name, state.reason);
      if (state.kind === 'claimed') {
        if (state.downloadFailure !== undefined) failures.set(name, state.downloadFailure);
        this.register(state, registered, failures);
      }
    }
    try {
      await this.load(delta, registered, failures, (loadCeiling ?? ceiling).signal);
    } finally {
      loadCeiling?.dispose();
    }
    if (this.closed) return this.settledState();
    this.warn(failures);
    this.snapshot = this.takeSnapshot();
    return this.snapshot;
  }

  /**
   * Acquire the claimed URL faces the library fetches itself (see
   * `face-download.ts`), replacing each by its bytes. A face that cannot be
   * fetched is dropped and its family warned; the family stays claimed (a
   * claimed face that fails never promotes another source or substitute).
   */
  private async acquire(resolved: Map<string, FamilyState>, ceiling: ReturnType<typeof ceilingSignal>): Promise<boolean> {
    type Claimed = Extract<FamilyState, { kind: 'claimed' }>;
    const omitted: [string, Claimed][] = [];
    const bridged = new Map<number, [string, Claimed][]>();
    for (const entry of resolved) {
      const state = entry[1];
      if (state.kind !== 'claimed' || !state.declarations.some(isDownloadableFace)) continue;
      const source = this.sources[state.sourceIndex];
      if (source.kind !== 'font-resolver') {
        if (source.credentials === 'omit') omitted.push([entry[0], state]);
      } else if (this.acquireResolverFaces) {
        bridged.set(state.sourceIndex, [...bridged.get(state.sourceIndex) ?? [], [entry[0], state]]);
      }
    }
    if (!omitted.length && !bridged.size) return false;
    const deadline = Date.now() + ceiling.remaining();
    const context = { signal: ceiling.signal, cancel: this.controller.signal, deadline, holder: this };
    const urlsOf = (state: Claimed) => state.declarations.filter(isDownloadableFace).map(({ url }) => url as string);
    const batches: Promise<[readonly [string, Claimed][], ReadonlyMap<string, FaceAcquisition>]>[] = [];
    if (omitted.length) {
      batches.push(acquireFaceUrls(omitted.flatMap(([, state]) => urlsOf(state)), { credentials: 'omit', ...context })
        .then((outcomes) => [omitted, outcomes]));
    }
    const acquire = this.acquireResolverFaces;
    for (const [sourceIndex, claimed] of bridged) {
      if (!acquire) break;
      const kept = new Map<string, FaceAcquisition>();
      const families = claimed.flatMap(([, state]) => {
        const urls = urlsOf(state).filter((url) => {
          const data = this.acquiredFaces.get(`${sourceIndex}\u0000${url}`);
          if (data) kept.set(url, { data });
          return !data;
        });
        return urls.length ? [{ family: state.name, urls }] : [];
      });
      if (!families.length) { batches.push(Promise.resolve([claimed, kept])); continue; }
      const request: ResolverFaceRequest = { source: this.sources[sourceIndex] as FontResolverSource, sourceIndex, families };
      batches.push(acquire(request, context).then((outcomes) => {
        for (const [url, outcome] of outcomes) {
          if (!('data' in outcome)) continue;
          const key = `${sourceIndex}\u0000${url}`;
          if (this.closed || this.acquiredFaces.has(key)) continue;
          if (this.ownsAcquiredFaces) {
            chargeRetainedBytes(outcome.data.byteLength);
            this.acquiredFaceBytes += outcome.data.byteLength;
          }
          this.acquiredFaces.set(key, outcome.data);
        }
        return [claimed, new Map([...kept, ...outcomes])];
      }));
    }
    const settled = await Promise.all(batches);
    if (this.closed) throw createAbortError('Font source session released');
    for (const [claimed, outcomes] of settled) {
      for (const [key, state] of claimed) {
        let reason: string | undefined;
        const declarations = state.declarations.flatMap((declaration) => {
          if (!isDownloadableFace(declaration)) return [declaration];
          const outcome = outcomes.get(declaration.url as string) ?? { reason: 'font download failed' };
          if ('data' in outcome) return [withAcquiredBytes(declaration, outcome.data)];
          reason ??= outcome.reason;
          return [];
        });
        resolved.set(key, { ...state, declarations, ...(reason === undefined ? {} : { downloadFailure: reason }) });
      }
    }
    return bridged.size > 0;
  }

  /** Merge demanded scalars; returns the newly demanded ones, or 'all' when
   * the session just became all-scalar (explicitly, or because the cumulative
   * set outgrew `HARD_MAX_FONT_DEMAND_SCALARS`). */
  private addScalars(scalars: 'all' | readonly number[]): 'all' | readonly number[] {
    if (this.all) return [];
    if (scalars !== 'all') {
      const delta: number[] = [];
      for (const scalar of scalars) if (!this.points.has(scalar)) { this.points.add(scalar); delta.push(scalar); }
      if (this.points.size <= HARD_MAX_FONT_DEMAND_SCALARS) return delta;
    }
    this.all = true;
    this.points.clear();
    return 'all';
  }

  /** Walk the sources in order for the fresh families (claim once). */
  private async resolve(fresh: NormalizedDemand['names']): Promise<Map<string, FamilyState>> {
    const result = new Map<string, FamilyState>();
    let pending = fresh;
    const ceiling = ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS);
    try {
      for (let sourceIndex = 0; sourceIndex < this.sources.length && pending.length; sourceIndex++) {
        if (this.closed) throw createAbortError('Font source session released');
        let answers: ReadonlyMap<string, FamilyResolution>;
        const source = this.sources[sourceIndex];
        // A runner that owns the deadline (the worker bridge: main times the
        // call out for the whole document) is not abandoned at this realm's
        // ceiling; only a release aborts the wait.
        const sourceSignal = source.kind === 'font-resolver' && this.runnerOwnsCeiling
          ? this.controller.signal : ceiling.signal;
        try {
          answers = await this.resolveWithSource(sourceIndex, pending, sourceSignal, ceiling.remaining);
        } catch (error) {
          const timedOut = isAbortError(error) && ceiling.expired() && sourceSignal === ceiling.signal;
          if (source.kind !== 'font-resolver') {
            if (!timedOut) throw error;
            for (const { name } of pending) result.set(fontFamilyKey(name), { kind: 'failed', name, reason: 'timed out' });
            pending = [];
            break;
          }
          // A release is no outcome. A timeout or any other error concludes
          // the batch per family (resolver-failures.ts `concludeFamilies`):
          // a family this document already decided (another requester's
          // resolution held in the cache, or a recorded failure) keeps it;
          // only undecided families fail, recorded for the whole document.
          if (isAbortError(error) && !timedOut) throw error;
          const cache = this.caches[sourceIndex];
          answers = this.claimFromCandidates(source, pending, concludeFamilies(this.resolverFailures, cache, pending,
            timedOut ? 'timed out' : error instanceof Error ? error.message : String(error), cache));
          // A ceiling breach still rejects the ensure.
          if (!timedOut) throw error;
        }
        const next: typeof pending[number][] = [];
        for (const entry of pending) {
          const key = fontFamilyKey(entry.name);
          const answer = answers.get(key) ?? { kind: 'absent' };
          if (answer.kind === 'absent') next.push(entry);
          else if (answer.kind === 'failed') result.set(key, { kind: 'failed', name: entry.name, reason: answer.reason });
          else {
            result.set(key, {
              kind: 'claimed', name: entry.name, sourceIndex, served: answer.served,
              ...(answer.substitution ? { substitution: answer.substitution } : {}),
              declarations: answer.declarations, faces: [],
            });
          }
        }
        pending = next;
      }
    } finally {
      ceiling.dispose();
    }
    return result;
  }

  /** A resolver answers with candidates; claim each family under its scope. */
  private claimFromCandidates(
    source: FontSource,
    families: NormalizedDemand['names'],
    answers: ReadonlyMap<string, FamilyResolution>,
  ): ReadonlyMap<string, FamilyResolution> {
    const claims = new Map(answers);
    for (const { name, scripts } of families) {
      const key = fontFamilyKey(name);
      const answer = answers.get(key);
      if (answer) claims.set(key, claimResolverCandidates(answer, name, source.substitutes, scripts));
    }
    return claims;
  }

  private async resolveWithSource(
    sourceIndex: number,
    families: NormalizedDemand['names'],
    signal: AbortSignal,
    remainingMs: () => number,
  ): Promise<ReadonlyMap<string, FamilyResolution>> {
    const source = this.sources[sourceIndex];
    const cache = this.caches[sourceIndex];
    const names = families.map(({ name }) => name);
    const scripts = new Map(families.map(({ name, scripts: proven }) => [fontFamilyKey(name), proven]));
    const scriptsOf = (name: string): FontSubstituteScripts => scripts.get(fontFamilyKey(name)) ?? 'all';
    switch (source.kind) {
      case 'font-css-api':
      case 'font-css':
        return resolveCssSource(source, cache, names, signal, scriptsOf);
      case 'font-files':
        return resolveFontFiles(source, cache, names, scriptsOf);
      case 'font-resolver':
        return this.claimFromCandidates(source, families,
          await this.resolveWithResolver(source, sourceIndex, cache, families, signal, remainingMs));
    }
  }

  /** Resolve through the realm's shared resolution cache (one application
   * call per family across every requester; `resolver-resolution.ts`). */
  private resolveWithResolver(
    source: FontResolverSource,
    sourceIndex: number,
    cache: SourceCache,
    families: NormalizedDemand['names'],
    signal: AbortSignal,
    remainingMs: () => number,
  ): Promise<ReadonlyMap<string, FamilyResolution>> {
    return resolveFamiliesThroughCache({
      source, cache, runner: this.runResolver, families,
      signal, holder: cache, sourceIndex, remainingMs,
      ...(this.resolverFailures ? { failures: this.resolverFailures } : {}),
    });
  }

  /** Admit and register one claimed family's declarations, in order. */
  private register(state: Extract<FamilyState, { kind: 'claimed' }>, registered: Set<OwnedFace>, failures: Map<string, string>): void {
    const set = this.fontSet as FontFaceSet;
    const sourceId = this.caches[state.sourceIndex].id;
    // Registry signature (design §4.4): source identity × family × every
    // descriptor × content, so two origins never share a face and two
    // documents on one source share one face per FontFaceSet.
    const signed = state.declarations.map((declaration) => {
      const { data, url = '' } = declaration;
      const dataUrl = data === undefined && /^\s*data:/i.test(url);
      const signature = [
        `src:${sourceId}`,
        fontFamilyKey(declaration.family),
        declaration.weight ?? '',
        declaration.style ?? '',
        declaration.stretch ?? '',
        declaration.unicodeRange ?? '',
        data !== undefined ? fontContentKey(data) : dataUrl ? fontContentKey(url) : url,
      ].join('|');
      return { declaration, signature, data, url, dataUrl };
    });
    // A face registered later in a FontFaceSet takes precedence over an
    // earlier one with the same descriptors (CSS Fonts 4 §4.10 / Font Loading
    // set order), so an identical declaration repeated in the list keeps the
    // position of its LAST occurrence: `A, B, A` registers `B, A`.
    const lastIndex = new Map(signed.map(({ signature }, index) => [signature, index]));
    for (const [index, { declaration, signature, data, url, dataUrl }] of signed.entries()) {
      if (lastIndex.get(signature) !== index) continue;
      const existing = this.owned.get(signature);
      if (existing) {
        existing.families.add(state.name);
        state.faces.push(existing);
        continue;
      }
      if (this.owned.size >= HARD_MAX_FONT_SOURCE_DECLARATIONS) {
        throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, this.owned.size + 1);
      }
      const bytes = data !== undefined ? data.byteLength : dataUrl ? dataUrlDecodedByteLength(url) ?? 0 : 0;
      chargeBytes(set, signature, bytes);
      let face: FontFace;
      try {
        ({ face } = retainFace(signature, set, () => {
          const descriptors: FontFaceDescriptors = {};
          if (declaration.weight) descriptors.weight = declaration.weight;
          if (declaration.style) descriptors.style = declaration.style;
          if (declaration.stretch) descriptors.stretch = declaration.stretch;
          if (declaration.unicodeRange) descriptors.unicodeRange = declaration.unicodeRange;
          const created = new FontFace(declaration.family, data ?? cssUrl(url), descriptors);
          try { set.add(created); }
          catch (error) { set.delete(created); throw error; }
          return created;
        }));
      } catch (error) {
        refundBytes(set, signature);
        failures.set(state.name, error instanceof Error ? error.message : String(error));
        continue;
      }
      const entry: OwnedFace = {
        signature, face, declaration, sourceIndex: state.sourceIndex,
        ranges: parseUnicodeRange(declaration.unicodeRange),
        charged: bytes > 0,
        families: new Set([state.name]),
      };
      this.owned.set(signature, entry);
      state.faces.push(entry);
      registered.add(entry);
    }
  }

  /** Load faces whose ranges intersect the demand (lease semantics: a face
   * registered now is checked against all scalars demanded so far, an older
   * unloaded face only against the delta). */
  private async load(
    delta: 'all' | readonly number[],
    registered: Set<OwnedFace>,
    failures: Map<string, string>,
    signal: AbortSignal,
  ): Promise<void> {
    const becameAll = delta === 'all';
    const deltaIndex = delta === 'all' ? [] : [...delta].sort((a, b) => a - b);
    const cumulative = registered.size && !this.all ? [...this.points].sort((a, b) => a - b) : deltaIndex;
    const waits: OwnedFace[] = [];
    for (const entry of this.owned.values()) {
      if (entry.readiness) { waits.push(entry); continue; }
      const isNew = registered.has(entry);
      const eligible = (this.all && (becameAll || isNew)) || entry.ranges === null
        || rangeIntersects(entry.ranges, isNew ? cumulative : deltaIndex);
      if (!eligible) continue;
      // load() is idempotent and joins a shared face's in-flight load
      // (CSS Font Loading §2.2), so a reused registration reports its own result.
      entry.readiness = Promise.resolve()
        .then(() => (this.closed ? undefined : entry.face.load()))
        .then(() => { entry.outcome = true; }, () => { entry.outcome = false; });
      waits.push(entry);
    }
    if (!waits.length) return;
    await until(Promise.all(waits.map((entry) => entry.readiness)), signal);
    if (this.closed) return;
    for (const entry of waits) {
      if (entry.outcome === true) continue;
      const reason = entry.outcome === false ? 'font load failed' : 'timed out';
      for (const family of entry.families) if (!failures.has(family)) failures.set(family, reason);
    }
  }

  private warn(failures: Map<string, string>): void {
    const fresh = [...failures].filter(([family]) => !this.warned.has(fontFamilyKey(family)));
    if (!fresh.length) return;
    for (const [family] of fresh) this.warned.add(fontFamilyKey(family));
    console.warn(`[ooxml] failed to preload web font(s): ${fresh.map(([family, reason]) => `${family} (${reason})`)
      .join(', ')}; falling back to system fonts (text may shift or differ).`);
  }

  private takeSnapshot(): LoadedFontSources {
    const faces: LoadedSourceFace[] = [];
    for (const entry of this.owned.values()) {
      if (entry.face.status !== 'loaded') continue;
      const { declaration } = entry;
      faces.push(Object.freeze({
        family: declaration.family,
        weight: declaration.weight ?? 'normal',
        style: declaration.style ?? 'normal',
        stretch: declaration.stretch ?? 'normal',
        unicodeRange: declaration.unicodeRange ?? 'U+0-10FFFF',
        sourceIndex: entry.sourceIndex,
        face: entry.face,
      }));
    }
    const substitutions = new Map<string, { readonly family: string; readonly script?: FontSubstituteScript }>();
    for (const [key, state] of this.families) {
      if (state.kind !== 'claimed' || !state.substitution) continue;
      if (!state.faces.some((entry) => entry.face.status === 'loaded')) continue;
      const script = state.substitution.script;
      substitutions.set(key, Object.freeze({ family: state.served, ...(script ? { script } : {}) }));
    }
    return Object.freeze({
      faces: Object.freeze(faces),
      substitutions,
      regularAliases: loadedRegularAliases(faces.map((face) => ({ ...face, status: 'loaded' as const }))),
    });
  }
}

const releasedRunner: FontResolverRunner = () => {
  throw createAbortError('Font source session released');
};
