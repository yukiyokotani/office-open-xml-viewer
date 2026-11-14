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
 *    throws, a run-time limit breach, the 15 s ceiling) also makes the
 *    family final: it keeps its host fallback and no later source is
 *    consulted, so a failing mirror never falls through to another origin
 *    (#1497 R5) and first paint is deterministic. Only a source that does not
 *    serve the family (an allow-list miss, a family absent from a stylesheet
 *    or file list) passes it on;
 * 2. **admits** the claimed declarations against the generated ceilings
 *    (families and declarations per session). A family whose demand or
 *    declarations would breach a ceiling fails alone, with a warning, like
 *    any other failed family: run-time limits are judged per family, never
 *    for the whole document (factory-time validation still throws);
 * 3. **acquires and registers** the declarations of a claimed family by
 *    overlap component, when the demand first intersects it. Every http(s) /
 *    blob: `url` face, whatever the source kind, is fetched by the library
 *    (`face-download.ts`) only then, one request per needed
 *    `unicode-range` subset, with the source's credentials (a resolver's
 *    through its acquirer: the worker bridge in a worker, this realm
 *    otherwise), and is registered from its bytes in the realm's FontFaceSet
 *    under this session's private alias of the family
 *    ({@link LoadedFontSources.families}), never under its real name, once,
 *    whichever subsets are demanded first. A component holds every
 *    declaration whose `unicode-range` overlaps another's, so overlapping
 *    precedence is independent of the order demand arrives in. A face that
 *    cannot be fetched (network, status, deadline, size or retained-byte
 *    ceiling) is dropped and its family warned;
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
 * same source identity (`source-cache.ts`), and downloaded bytes across
 * sessions of the same realm (`face-download.ts`). {@link
 * FontSourceSession.release} returns every reference the session took,
 * exactly once.
 *
 * Document scope (library policy). The FontFaceSet is shared by every
 * document of a realm, and CSS font matching selects among all faces of a
 * family name: a face registered under its real family would paint (and size)
 * every other open document that names that family, including one whose own
 * sources serve different faces under the same name. Each session therefore
 * registers its faces under private family names of its own (one per served
 * family), and the formats map a served family to that alias in every font
 * stack they build from the snapshot (`scopeServedFontFamilies`). A document's
 * painting and line metrics then depend only on its own session: no outcome,
 * registration or release of another document can change them. The cost is
 * one registration (and one browser decode) per open document of a face that
 * several documents share; the downloaded bytes are still shared.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import type { OoxmlFontSourceQuota } from '../errors/ooxml-error.js';
import { createAbortError } from '../internal/abort.js';
import {
  HARD_MAX_FONT_DEMAND_SCALARS,
  HARD_MAX_FONT_SOURCE_DECLARATIONS,
  HARD_MAX_FONT_SOURCE_FAMILIES,
} from '../worker/resource-policy.generated.js';
import { resolveCssSource } from './css-source.js';
import { parseUnicodeRange } from './unicode-range.js';
import {
  FONT_DEADLINE_REASON,
  acquireFaceUrls,
  inRealmResolverFaceAcquirer,
  isDownloadableFace,
  releaseAcquiredFaces,
  withAcquiredBytes,
  type FaceAcquisition,
  type ResolverFaceAcquirer,
  type ResolverFaceRequest,
} from './face-download.js';
import { resolveFontFiles } from './files-source.js';
import type { OpenTypeLineMetrics } from './open-type-metrics.js';
import { DecodeQueue } from './decode-queue.js';
import { concludeFamilies } from './resolver-failures.js';
import {
  claimResolverCandidates,
  resolveFamiliesThroughCache,
  runResolverInRealm,
  substituteScriptsOf,
  type FontResolverFailures,
  type FontResolverRunner,
} from './resolver-resolution.js';
import { activeFontSet, FONT_LOAD_CEILING_MS } from './font-registry.js';
import {
  acquireSourceCache,
  addRetainedReclaimer,
  chargeReclaimableBytes,
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
  type FontCssApiSource,
  type FontResolverSource,
  type FontSource,
  type FontSourceDescriptor,
  type FontSourceCredentials,
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
  /** OpenType tables of the exact bytes this face was registered from, when
   * the session holds them (`data` faces and every library-acquired
   * http(s) / blob: download), the session reads tables
   * ({@link FontSourceSessionOptions.readMetricTables}) and the parser reads them: TrueType/OpenType, WOFF, and WOFF2 where the engine
   * inflates Brotli natively. A variable font carries its `variation` data
   * for {@link instanceOpenTypeLineMetrics}. Formats project these tables
   * with their own line rules, exactly as for a document-embedded face. */
  readonly resourceMetrics?: OpenTypeLineMetrics;
  /** When `resourceMetrics` is absent: the source this face was registered
   * from (the owned bytes, or the `url()` of an inline data: URL), so a
   * format can register it again alone under a private family and measure
   * it in the DOM without a request. Absent for a face the tables were read
   * from. */
  readonly probeSource?: string | ArrayBuffer;
  /** The supplying source guarantees that its `unicode-range` subsets of one
   * served family with equal descriptors share one font file's vertical
   * metric tables ({@link FontCssApiSource.sharedSubsetMetricFamilies}); present
   * only with {@link probeSource}. */
  readonly sharedSubsetMetrics?: true;
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
  /** Lower-cased served family (`LoadedSourceFace.family`) → the private CSS
   * family this session registered its faces under, for every family with a
   * registered face. A font stack names the alias immediately before the
   * served family (`scopeServedFontFamilies`): the session's faces win
   * exactly where a face registered under the real name would, and nothing
   * another document registers can be selected for this one. */
  readonly families: ReadonlyMap<string, string>;
}

export type { FontResolverFailures, FontResolverRunner, ResolverFaceAcquirer, ResolverFaceRequest };

/**
 * The state of the `fontResolver` worker bridge's main-side host for one
 * document (`resolver-bridge-gate.ts`), as a session of that document sees it:
 * the render worker's (`WorkerFontResolverClient`) and, for XLSX, the viewer
 * realm's, which shares the document's resolver outcomes with the worker.
 */
export interface ResolverBridgeState {
  /** While the host is loading: settles once it is ready or unavailable.
   * Undefined otherwise. */
  readonly ready: () => Promise<void> | undefined;
  /** Once the host is unavailable: the reason every family of the
   * document's resolver sources fails with. */
  readonly unavailable: () => string | undefined;
}

export interface FontSourceSessionOptions {
  /** The `fonts` option value (factory results or a worker's copies); validated here. */
  readonly sources: readonly (FontSource | FontSourceDescriptor)[];
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
  /** The worker bridge's host state (see {@link ResolverBridgeState}). A
   * session with a resolver source starts each source walk, and its font
   * ceiling, only once the host is ready: no deadline of the document's
   * resolver calls runs while the host loads. Once the host is unavailable,
   * every family of a resolver source fails with its reason (host-font
   * fallback), whatever the realm cache holds, in every realm alike. */
  readonly resolverBridge?: ResolverBridgeState;
  /** Acquires the needed http(s) / blob: faces of resolver sources, as the
   * demand reaches them (see {@link ResolverFaceAcquirer}); defaults to this
   * realm's acquisition through `resolverFailures`. */
  readonly acquireResolverFaces?: ResolverFaceAcquirer;
  /** The acquirer hands over fresh buffers this session alone keeps (the
   * worker bridge's transferred copies), so the session charges each one
   * against the realm ceiling while it keeps it. Otherwise the realm's
   * acquisition cache owns and charges them. */
  readonly acquiredFacesOwnedBySession?: boolean;
  /** Reads the OpenType metric tables of each loaded face that has bytes
   * (`source-metric-tables.ts` `readFontResourceMetrics`), before the
   * snapshot: {@link LoadedSourceFace.resourceMetrics}. Only a format that
   * lays out with those tables passes it (DOCX). Absent, no face is inflated
   * or parsed, and every face carries its {@link LoadedSourceFace.probeSource}. */
  readonly readMetricTables?: (data: ArrayBuffer) => Promise<OpenTypeLineMetrics | null>;
}


/** Parsed tables per face buffer (null: unreadable). Downloaded bytes are
 * one buffer for every session of the realm that holds them
 * (`face-download.ts`), so documents sharing a face decode it once. Decoding
 * is bounded through {@link metricDecodes}. */
const faceResourceMetrics = new WeakMap<ArrayBuffer, OpenTypeLineMetrics | null>();

/** Serial of the sessions created in this realm: each session's private
 * family names ({@link LoadedFontSources.families}) carry its own serial, so
 * no two sessions of one realm ever share one. */
let sessionSerial = 0;

/** Metric-table decodes of this realm: a WOFF/WOFF2 decode inflates up to
 * `MAX_INFLATED_BYTES`, so at most two run at once; a released document's
 * queued decodes are dropped (`decode-queue.ts`). */
const metricDecodes = new DecodeQueue(2);

/** Read the metric tables of `data` once with `read`, on behalf of
 * `holder`; rejects with an `AbortError` once `signal` aborts. */
function prepareResourceMetrics(
  data: ArrayBuffer,
  read: (data: ArrayBuffer) => Promise<OpenTypeLineMetrics | null>,
  holder: object,
  signal: AbortSignal,
): Promise<void> {
  if (faceResourceMetrics.has(data)) return Promise.resolve();
  return metricDecodes.run(data, holder, () => decodeResourceMetrics(data, read), signal);
}

/** Decodes queued or running in this realm (tests, leak checks). */
export function pendingMetricDecodeCount(): number {
  return metricDecodes.size;
}

async function decodeResourceMetrics(
  data: ArrayBuffer,
  read: (data: ArrayBuffer) => Promise<OpenTypeLineMetrics | null>,
): Promise<void> {
  if (faceResourceMetrics.has(data)) return;
  let metrics: OpenTypeLineMetrics | null = null;
  try { metrics = await read(data); } catch { metrics = null; }
  faceResourceMetrics.set(data, metrics);
}

const EMPTY_SNAPSHOT: LoadedFontSources = Object.freeze({
  faces: Object.freeze([]),
  substitutions: new Map(),
  regularAliases: new Map(),
  families: new Map(),
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

/** A face's `font-style` when it is a plain keyword (omitted = `normal`);
 * undefined for an oblique angle range or anything else. */
function plainStyle({ style = 'normal' }: FontFaceDeclaration): string | undefined {
  const text = style.trim().toLowerCase();
  return text === 'normal' || text === 'italic' || text === 'oblique' ? text : undefined;
}

/** Whether two faces can be candidates for one style request: CSS Fonts 4
 * §5.2 narrows by stretch, style and weight before `unicode-range`. Only a
 * difference in plain style keywords (upright vs italic) separates them:
 * weights and stretches are kept together whatever their values, which
 * costs at most a few extra subsets of another weight, never precedence. */
function mayMatchTogether(a: FontFaceDeclaration, b: FontFaceDeclaration): boolean {
  const left = plainStyle(a), right = plainStyle(b);
  return left === undefined || right === undefined || left === right;
}

/**
 * Overlap components of a family's declarations: two declarations whose
 * `unicode-range`s overlap share a component, transitively, and a
 * declaration without a range shares one with every declaration that can be
 * a candidate for the same style request ({@link mayMatchTogether}). CSS font
 * matching orders only faces that can serve the same scalar for the same
 * request (CSS Fonts 4 §4.5, §5), so registering whole components, each in
 * declaration order, keeps every precedence the full declaration order
 * defines, whatever order components are needed in, while an unranged face
 * of one style never pulls in the subsets of another.
 */
function overlapComponents(declarations: readonly FontFaceDeclaration[]): number[] {
  const parent = declarations.map((_, index) => index);
  const find = (index: number): number => {
    while (parent[index] !== index) { parent[index] = parent[parent[index]!]!; index = parent[index]!; }
    return index;
  };
  const union = (a: number, b: number) => { const ra = find(a), rb = find(b); if (ra !== rb) parent[Math.max(ra, rb)] = Math.min(ra, rb); };
  const intervals: [number, number, number][] = [];
  declarations.forEach((declaration, index) => {
    const ranges = parseUnicodeRange(declaration.unicodeRange);
    if (ranges === null) {
      declarations.forEach((other, at) => { if (at !== index && mayMatchTogether(declaration, other)) union(index, at); });
      return;
    }
    for (const [lo, hi] of ranges) intervals.push([lo, hi, index]);
  });
  intervals.sort((a, b) => a[0] - b[0]);
  let reach = -1;
  let owner = -1;
  for (const [lo, hi, index] of intervals) {
    if (lo <= reach) union(owner, index);
    if (hi > reach) { reach = hi; owner = index; }
  }
  return declarations.map((_, index) => find(index));
}

/** One declaration's identity for the declaration ceiling: its family,
 * descriptors and source (URL, or the owned bytes' content key). */
function declarationKey(declaration: FontFaceDeclaration): string {
  return [
    fontFamilyKey(declaration.family), declaration.weight ?? '', declaration.style ?? '', declaration.stretch ?? '',
    declaration.unicodeRange ?? '', declaration.data !== undefined ? fontContentKey(declaration.data) : declaration.url ?? '',
  ].join('|');
}

/** {@link FontSourceSession}'s registration signature of one
 * declaration: source × family × every descriptor × content, so a face two
 * claimed families share is registered once. */
function faceSignature(sourceIndex: number, declaration: FontFaceDeclaration): string {
  const { data, url = '' } = declaration;
  const dataUrl = data === undefined && /^\s*data:/i.test(url);
  return [
    sourceIndex,
    fontFamilyKey(declaration.family),
    declaration.weight ?? '',
    declaration.style ?? '',
    declaration.stretch ?? '',
    declaration.unicodeRange ?? '',
    data !== undefined ? fontContentKey(data) : dataUrl ? fontContentKey(url) : url,
  ].join('|');
}

/** The batch a source's own downloads join: one per credentials mode and
 * download pool. */
function directGroup(source: Exclude<FontSourceDescriptor, FontResolverSource>): string {
  const parallel = source.kind === 'font-css-api' ? source.parallelFaceDownloads : undefined;
  return `${source.credentials ?? 'same-origin'}\u0000${parallel ?? ''}`;
}

/** The `FontFace` descriptors of one declaration. */
function faceDescriptors(declaration: FontFaceDeclaration): FontFaceDescriptors {
  const descriptors: FontFaceDescriptors = {};
  if (declaration.weight) descriptors.weight = declaration.weight;
  if (declaration.style) descriptors.style = declaration.style;
  if (declaration.stretch) descriptors.stretch = declaration.stretch;
  if (declaration.unicodeRange) descriptors.unicodeRange = declaration.unicodeRange;
  return descriptors;
}

/** Faces are no longer prepared ahead this close to the face-loading
 * deadline, so no added work can delay a download's deadline check there. */
const PREPARE_DEADLINE_MARGIN_MS = 1_000;

/** A CSS `url()` for a validated absolute URL (data: URLs are kept verbatim
 * by validation and may contain characters a CSS string must escape). */
function cssUrl(url: string): string {
  return `url("${url.replace(/[\\"\n\r\f]/g, (char) =>
    char === '\\' || char === '"' ? `\\${char}` : `\\${char.charCodeAt(0).toString(16)} `)}")`;
}

/** The one warning of an ensure: each failed family (or group) with its reason. */
function warnFailures(failures: readonly (readonly [string, string])[]): void {
  console.warn(`[ooxml] failed to preload web font(s): ${failures.map(([family, reason]) => `${family} (${reason})`)
    .join(', ')}; falling back to system fonts (text may shift or differ).`);
}

/** The warning reason of a run-time ceiling breach: the family fails alone. */
function limitReason(quota: OoxmlFontSourceQuota, limit: number, observed: number): string {
  return fontSourceLimitError(quota, limit, observed).message;
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

export { retainedFontSourceBytes };

// ---- session ------------------------------------------------------------------------

interface OwnedFace {
  readonly face: FontFace;
  readonly declaration: FontFaceDeclaration;
  readonly sourceIndex: number;
  /** null: no (or an unparseable) unicode-range, so the face covers all. */
  readonly ranges: readonly (readonly [number, number])[] | null;
  /** Decoded bytes charged to the realm ceiling while registered. */
  readonly bytes: number;
  /** Demanded families this face serves (for warnings). */
  readonly families: Set<string>;
  readiness?: Promise<void>;
  /** Settled load outcome; undefined while loading. */
  outcome?: boolean;
}

/** The demand-driven faces of a claimed family (every source kind): its
 * `url` faces are fetched when they are needed. Declarations are registered by
 * overlap component (see {@link overlapComponents}), so a face is acquired
 * and registered once, from its bytes, whichever subsets are demanded first,
 * and overlapping `unicode-range` precedence keeps declaration order. */
interface DemandFaces {
  /** Overlap component of each declaration. */
  readonly component: readonly number[];
  /** Per declaration: registered, or dropped (its download failed). */
  readonly settled: boolean[];
}

type FamilyState =
  | { readonly kind: 'claimed'; readonly name: string; readonly sourceIndex: number; readonly served: string;
    /** Current representation: a lazily fetched face is replaced by its
     * bytes once acquired. */
    readonly substitution?: FontSubstitution; readonly declarations: FontFaceDeclaration[];
    readonly faces: OwnedFace[];
    /** Set once the family is admitted (every claimed family of the session). */
    readonly demand?: DemandFaces }
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
  private sources: readonly FontSourceDescriptor[];
  private caches: readonly SourceCache[];
  private fontSet: FontFaceSet | null;
  private runResolver: FontResolverRunner;
  private resolverFailures: FontResolverFailures | undefined;
  private acquireResolverFaces: ResolverFaceAcquirer | undefined;
  private readonly ownsAcquiredFaces: boolean;
  private readonly readMetricTables: FontSourceSessionOptions['readMetricTables'];
  /** Resolver faces this session acquired, by `sourceIndex\0url`: a later
   * delta claiming the same face reuses the bytes instead of asking again. */
  private readonly acquiredFaces = new Map<string, ArrayBuffer>();
  private acquiredFaceBytes = 0;
  private readonly runnerOwnsCeiling: boolean;
  private resolverBridge: ResolverBridgeState | undefined;
  private readonly controller = new AbortController();
  private detachSignal: () => void = () => undefined;
  private closed = false;
  private abortedByCaller = false;
  private failure: unknown = undefined;
  private queue: Promise<unknown> = Promise.resolve();
  private readonly families = new Map<string, FamilyState>();
  /** Registered faces, by {@link register}'s signature. */
  private readonly owned = new Map<string, OwnedFace>();
  /** Served family key → this session's private family for it. */
  private readonly aliases = new Map<string, string>();
  /** Faces built from downloaded bytes as each download completed (while the
   * rest of its batch still downloads), by {@link register}'s signature,
   * under a placeholder family; {@link register} names and adds them in its
   * own order. Emptied after each registration pass. */
  private readonly prepared = new Map<string, { readonly face: FontFace; readonly bytes: number }>();
  /** Unregisters {@link dropPrepared} from the realm's reclaimers. */
  private stopReclaim: (() => void) | undefined;
  private readonly serial = ++sessionSerial;
  /** Every declaration this session claimed (by {@link declarationKey}). */
  private readonly claimedDeclarations = new Set<string>();
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
    this.resolverBridge = this.sources.some((source) => source.kind === 'font-resolver') ? options.resolverBridge : undefined;
    // A resolver's URL faces are always fetched by the library, per needed
    // subset: by the page for a worker (the bridge), and in this realm
    // otherwise, so both modes send the same requests and hold the bytes.
    this.acquireResolverFaces = options.acquireResolverFaces ?? inRealmResolverFaceAcquirer(options.resolverFailures);
    this.ownsAcquiredFaces = options.acquiredFacesOwnedBySession === true;
    this.readMetricTables = options.readMetricTables;
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
   * of claimed families). Resolves with the snapshot; a family that fails
   * (including a run-time limit breach) keeps its host fallback with a
   * warning. Rejects with an `AbortError` once the caller's signal aborted.
   * After {@link release} it resolves with an empty snapshot.
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
    this.dropPrepared();
    for (const entry of this.owned.values()) {
      this.fontSet?.delete(entry.face);
      refundRetainedBytes(entry.bytes);
    }
    this.owned.clear();
    this.aliases.clear();
    this.claimedDeclarations.clear();
    this.families.clear();
    this.points.clear();
    for (const cache of this.caches) cache.release();
    releaseAcquiredFaces(this);
    metricDecodes.release(this);
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
    this.resolverBridge = undefined;
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
      const failures = new Map<string, string>();
      const overflow: string[] = [];
      // Families beyond the ceiling are never recorded, so `families` holds
      // exactly the ones counted against it.
      let tracked = this.families.size;
      for (const entry of demand.names) {
        const key = fontFamilyKey(entry.name);
        const prior = this.families.get(key);
        if (!prior) {
          // The family ceiling is document demand, judged per family: the
          // families beyond it fail (final, warned), the rest load.
          // Their names are not kept (that would grow with the document
          // past the ceiling): each delta naming one warns again.
          if (tracked >= HARD_MAX_FONT_SOURCE_FAMILIES) {
            overflow.push(entry.name);
            continue;
          }
          tracked++;
          fresh.push(entry);
          continue;
        }
        if (prior.kind === 'absent' && scopeWidens(prior.scripts, entry.scripts)) {
          fresh.push({ ...entry, scripts: widenScope(prior.scripts, entry.scripts) });
        }
      }
      // One bounded warning for all of them (count and first names), not
      // recorded in `warned`: nothing about them is retained.
      if (overflow.length) {
        warnFailures([[`families over the limit: ${overflow.length} [${overflow.slice(0, 8).join(', ')}${overflow.length > 8 ? ', …' : ''}]`,
          limitReason('families', HARD_MAX_FONT_SOURCE_FAMILIES, tracked + overflow.length)]]);
      }
      const delta = this.addScalars(demand.scalars);
      const resolved = fresh.length ? await this.resolve(fresh) : new Map<string, FamilyState>();
      if (this.closed) return this.settledState();
      // One face-loading ceiling covers acquisition and loading, as the
      // browser's own fetch inside FontFace.load() would.
      const ceiling = ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS);
      try {
        return await this.acquireRegisterLoad(fresh, resolved, delta, ceiling, failures);
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
    failures: Map<string, string>,
  ): Promise<LoadedFontSources> {
    for (const { name, scripts } of fresh) {
      const key = fontFamilyKey(name);
      let state = resolved.get(key) ?? { kind: 'absent', name, scripts };
      if (state.kind === 'claimed') {
        // The declaration ceiling covers every claimed declaration, acquired
        // now or on demand later, so a breach is decided at claim whatever
        // the text: the family that would cross it fails alone.
        const { sourceIndex } = state;
        const keys = new Set(state.declarations.map((declaration) => `${sourceIndex}\u0000${declarationKey(declaration)}`));
        let size = this.claimedDeclarations.size;
        for (const declaration of keys) if (!this.claimedDeclarations.has(declaration)) size++;
        if (size > HARD_MAX_FONT_SOURCE_DECLARATIONS) {
          state = { kind: 'failed', name, reason: limitReason('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, size) };
        } else {
          for (const declaration of keys) this.claimedDeclarations.add(declaration);
          state = { ...state, demand: { component: overlapComponents(state.declarations), settled: state.declarations.map(() => false) } };
        }
      }
      this.families.set(key, state);
      if (state.kind === 'failed') failures.set(name, state.reason);
    }
    const registered = new Set<OwnedFace>();
    const bridged = await this.registerDemanded(registered, failures, ceiling);
    if (this.closed) return this.settledState();
    // A resolver acquisition answers with the page's verdict, which may arrive
    // after this realm's ceiling (main owns that deadline in worker mode): the
    // faces it delivered still get a bounded wait for decoding, so this realm
    // paints what the page decided, as another realm of the document does.
    // By the absolute deadline, not the ceiling's timer, which may not have
    // run yet when a busy realm receives the reply.
    const loadCeiling = bridged && ceiling.remaining() === 0
      ? ceilingSignal(this.controller.signal, FONT_LOAD_CEILING_MS) : undefined;
    try {
      await this.load(delta, registered, failures, (loadCeiling ?? ceiling).signal);
    } finally {
      loadCeiling?.dispose();
    }
    if (this.closed) return this.settledState();
    const read = this.readMetricTables;
    if (read) {
      await Promise.all([...this.owned.values()].flatMap((entry) =>
        entry.face.status === 'loaded' && entry.declaration.data !== undefined
          ? [prepareResourceMetrics(entry.declaration.data, read, this, this.controller.signal)] : []));
    }
    if (this.closed) return this.settledState();
    this.warn(failures);
    this.snapshot = this.takeSnapshot();
    return this.snapshot;
  }

  /**
   * Acquire and register the faces the demand so far needs: every not yet
   * registered overlap component of a claimed family that holds a declaration
   * intersecting the cumulative scalars (or every one, for 'all'; a
   * declaration without a range always). One path for every source kind: a
   * `url` face is fetched by the library (CORS, one request per face), with
   * the source's `credentials` for a stylesheet or file source and through
   * the resolver acquirer for a {@link fontResolver} answer (the page fetches
   * it, in both modes, with the CSS font-loading credentials), and is
   * registered from those bytes, so its tables can be read and the file is
   * downloaded once. Downloads are shared with other sessions of this realm
   * while any of them holds them. A face that cannot be fetched is dropped
   * and its family warned; the family stays claimed (a claimed face that
   * fails never promotes another source or substitute). Returns whether a
   * resolver acquisition ran (its verdict may postdate this realm's ceiling).
   */
  private async registerDemanded(
    registered: Set<OwnedFace>,
    failures: Map<string, string>,
    ceiling: ReturnType<typeof ceilingSignal>,
  ): Promise<boolean> {
    type Claimed = Extract<FamilyState, { kind: 'claimed' }>;
    const sorted = this.all ? [] : [...this.points].sort((a, b) => a - b);
    const work: { state: Claimed; demand: DemandFaces; indices: number[] }[] = [];
    for (const state of this.families.values()) {
      if (state.kind !== 'claimed' || !state.demand) continue;
      const { component, settled } = state.demand;
      const needed = new Set<number>();
      state.declarations.forEach((declaration, index) => {
        if (settled[index]) return;
        const ranges = parseUnicodeRange(declaration.unicodeRange);
        if (this.all || ranges === null || rangeIntersects(ranges, sorted)) needed.add(component[index]!);
      });
      if (!needed.size) continue;
      const indices = state.declarations.flatMap((_, index) => !settled[index] && needed.has(component[index]!) ? [index] : []);
      work.push({ state, demand: state.demand, indices });
    }
    if (!work.length) return false;
    // The needed URLs, per acquisition path: a credentials mode for the
    // library's own requests, or a resolver source (its acquirer).
    const direct = new Map<string, { credentials: FontSourceCredentials; parallel: number | undefined; urls: Set<string> }>();
    // Faces each URL's bytes will register as (their signatures need the
    // bytes), prepared as soon as those bytes arrive.
    const pending = new Map<string, { sourceIndex: number; declaration: FontFaceDeclaration }[]>();
    const viaResolver = new Map<number, Map<string, string[]>>();
    for (const { state, indices } of work) {
      const source = this.sources[state.sourceIndex];
      for (const index of indices) {
        const declaration = state.declarations[index]!;
        if (!isDownloadableFace(declaration)) continue;
        const url = declaration.url as string;
        if (source.kind === 'font-resolver') {
          let families = viaResolver.get(state.sourceIndex);
          if (!families) { families = new Map(); viaResolver.set(state.sourceIndex, families); }
          families.set(state.name, [...families.get(state.name) ?? [], url]);
        } else {
          const credentials = source.credentials ?? 'same-origin';
          const parallel = source.kind === 'font-css-api' ? source.parallelFaceDownloads : undefined;
          const group = directGroup(source);
          let batch = direct.get(group);
          if (!batch) { batch = { credentials, parallel, urls: new Set() }; direct.set(group, batch); }
          batch.urls.add(url);
          let faces = pending.get(url);
          if (!faces) { faces = []; pending.set(url, faces); }
          faces.push({ sourceIndex: state.sourceIndex, declaration });
        }
      }
    }
    const deadline = Date.now() + ceiling.remaining();
    const outcomes = new Map<string, ReadonlyMap<string, FaceAcquisition>>();
    const batches: Promise<void>[] = [];
    // Building a face from bytes decodes them synchronously; doing it as
    // each download lands overlaps that work with the downloads still in
    // flight instead of running it all after the last one.
    // Preparation runs in a task of its own after the downloads that landed
    // (their deadline checks come first), and no face is built within
    // PREPARE_DEADLINE_MARGIN_MS of the deadline (checked before each), so
    // preparing never delays a download's deadline check past the deadline.
    const landed: [string, ArrayBuffer][] = [];
    let flush: ReturnType<typeof setTimeout> | undefined;
    const prepareLanded = () => {
      flush = undefined;
      for (const [url, data] of landed.splice(0)) if (!prepare(url, data)) break;
    };
    const onAcquired = (url: string, data: ArrayBuffer) => {
      if (this.closed) return;
      landed.push([url, data]);
      flush ??= setTimeout(prepareLanded, 0);
    };
    /** False once the margin is reached: nothing more is prepared. */
    const prepare = (url: string, data: ArrayBuffer): boolean => {
      for (const { sourceIndex, declaration } of pending.get(url) ?? []) {
        // Checked before any work for a face (hashing included).
        if (this.closed || Date.now() + PREPARE_DEADLINE_MARGIN_MS > deadline) return false;
        const acquired = withAcquiredBytes(declaration, data);
        const signature = faceSignature(sourceIndex, acquired);
        if (this.owned.has(signature) || this.prepared.has(signature)) continue;
        // The face's copy of the bytes is charged while it is prepared, only
        // from room the realm has: any charge that would be refused first
        // drops every prepared face (dropPrepared), so preparing never
        // decides a download or registration, and never passes the ceiling.
        if (!chargeReclaimableBytes(data.byteLength)) continue;
        try {
          const face = new FontFace(`ooxml-font-source-${this.serial}-prepared`, data, faceDescriptors(acquired));
          this.prepared.set(signature, { face, bytes: data.byteLength });
          this.stopReclaim ??= addRetainedReclaimer(() => this.dropPrepared());
        } catch {
          // Left to register(), which builds it again and records the error.
          refundRetainedBytes(data.byteLength);
        }
      }
      return true;
    };
    for (const [group, { credentials, parallel, urls }] of direct) {
      batches.push(acquireFaceUrls(urls, {
        credentials, signal: ceiling.signal, deadline, holder: this, onAcquired,
        ...(parallel === undefined ? {} : { parallel }),
      }).then((acquired) => { outcomes.set(group, acquired); }));
    }
    const acquire = this.acquireResolverFaces;
    for (const [sourceIndex, byFamily] of viaResolver) {
      if (!acquire) break;
      batches.push(this.acquireResolverUrls(acquire, sourceIndex, byFamily, ceiling, deadline)
        .then((acquired) => { outcomes.set(`resolver:${sourceIndex}`, acquired); }));
    }
    try {
      await Promise.all(batches);
      if (this.closed) throw createAbortError('Font source session released');
      // Prepared faces stay charged (and reclaimable) until registration
      // takes them over in declaration order; see register().
      this.registerAcquired(work, outcomes, registered, failures);
    } finally {
      if (flush !== undefined) clearTimeout(flush);
      landed.length = 0;
      // Prepared faces no registration took (a missed deadline, a face
      // another family registered first, a released session) are dropped.
      this.dropPrepared();
    }
    return viaResolver.size > 0 && acquire !== undefined;
  }

  /** Drop every prepared face and return its charge. */
  private dropPrepared(): void {
    for (const entry of this.prepared.values()) refundRetainedBytes(entry.bytes);
    this.prepared.clear();
    this.stopReclaim?.();
    this.stopReclaim = undefined;
  }

  /** Register the acquired faces of each work item, in declaration order. */
  private registerAcquired(
    work: readonly { state: Extract<FamilyState, { kind: 'claimed' }>; demand: DemandFaces; indices: readonly number[] }[],
    outcomes: ReadonlyMap<string, ReadonlyMap<string, FaceAcquisition>>,
    registered: Set<OwnedFace>,
    failures: Map<string, string>,
  ): void {
    for (const { state, demand, indices } of work) {
      const source = this.sources[state.sourceIndex];
      const group = outcomes.get(source.kind === 'font-resolver' ? `resolver:${state.sourceIndex}` : directGroup(source));
      const ready: number[] = [];
      for (const index of indices) {
        demand.settled[index] = true;
        const declaration = state.declarations[index]!;
        if (!isDownloadableFace(declaration)) { ready.push(index); continue; }
        const outcome = group?.get(declaration.url as string) ?? { reason: 'font download failed' };
        if ('data' in outcome) {
          state.declarations[index] = withAcquiredBytes(declaration, outcome.data);
          ready.push(index);
        } else if (!failures.has(state.name)) {
          failures.set(state.name, outcome.reason);
        }
      }
      this.register(state, registered, failures, ready);
    }
  }

  /** Acquire the needed URL faces of one resolver source's claimed families
   * through the resolver acquirer, reusing bytes this session already holds
   * (two families claiming one face in different deltas). */
  private async acquireResolverUrls(
    acquire: ResolverFaceAcquirer,
    sourceIndex: number,
    byFamily: ReadonlyMap<string, readonly string[]>,
    ceiling: ReturnType<typeof ceilingSignal>,
    deadline: number,
  ): Promise<ReadonlyMap<string, FaceAcquisition>> {
    const outcomes = new Map<string, FaceAcquisition>();
    const families: { family: string; urls: string[] }[] = [];
    for (const [family, list] of byFamily) {
      const urls = [...new Set(list)].filter((url) => {
        const data = this.acquiredFaces.get(`${sourceIndex}\u0000${url}`);
        if (data) outcomes.set(url, { data });
        return !data;
      });
      if (urls.length) families.push({ family, urls });
    }
    if (!families.length) return outcomes;
    const request: ResolverFaceRequest = { source: this.sources[sourceIndex] as FontResolverSource, sourceIndex, families };
    const acquired = await acquire(request, { signal: ceiling.signal, cancel: this.controller.signal, deadline, holder: this });
    for (const [url, outcome] of acquired) {
      const key = `${sourceIndex}\u0000${url}`;
      if (!('data' in outcome) || this.closed) { outcomes.set(url, outcome); continue; }
      const kept = this.acquiredFaces.get(key);
      if (kept) { outcomes.set(url, { data: kept }); continue; }
      if (this.ownsAcquiredFaces) {
        // Fresh copies this session alone keeps: charged while it keeps them;
        // beyond the realm ceiling the face fails like any oversized download.
        try {
          chargeRetainedBytes(outcome.data.byteLength);
        } catch (error) {
          outcomes.set(url, { reason: error instanceof Error ? error.message : String(error) });
          continue;
        }
        this.acquiredFaceBytes += outcome.data.byteLength;
      }
      this.acquiredFaces.set(key, outcome.data);
      outcomes.set(url, outcome);
    }
    return outcomes;
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
    // The walk and its ceiling start once the bridge host can answer.
    const hostReady = this.resolverBridge?.ready();
    if (hostReady) {
      await until(hostReady, this.controller.signal);
      if (this.closed) throw createAbortError('Font source session released');
    }
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
          // A release is no outcome. A timeout or any other error is the
          // source's failure for the families still pending: final, warned.
          if (isAbortError(error) && !timedOut) throw error;
          const reason = timedOut ? 'timed out' : error instanceof Error ? error.message : String(error);
          if (source.kind !== 'font-resolver') {
            for (const { name } of pending) result.set(fontFamilyKey(name), { kind: 'failed', name, reason });
            pending = [];
            break;
          }
          // A resolver batch is concluded per family (resolver-failures.ts
          // `concludeFamilies`): a family this document already decided
          // (another requester's resolution held in the cache, or a recorded
          // failure) keeps it; only undecided families fail, recorded for the
          // whole document.
          const cache = this.caches[sourceIndex];
          answers = this.claimFromCandidates(source, pending, concludeFamilies(this.resolverFailures, cache, pending,
            reason, cache));
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
              // A copy: on-demand faces replace their entries with bytes, and the
              // answer is shared through the source cache.
              declarations: [...answer.declarations], faces: [],
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
    source: FontSourceDescriptor,
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
      case 'font-resolver': {
        const unavailable = this.resolverBridge?.unavailable();
        if (unavailable !== undefined) {
          return new Map(families.map(({ name }) => [fontFamilyKey(name), { kind: 'failed', reason: unavailable } satisfies FamilyResolution]));
        }
        return this.claimFromCandidates(source, families,
          await this.resolveWithResolver(source, sourceIndex, cache, families, signal, remainingMs));
      }
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

  /** Register the declarations `indices` (ascending) of one claimed family,
   * in order. */
  private register(
    state: Extract<FamilyState, { kind: 'claimed' }>,
    registered: Set<OwnedFace>,
    failures: Map<string, string>,
    indices: readonly number[],
  ): void {
    const set = this.fontSet as FontFaceSet;
    const chosen = indices.map((index) => state.declarations[index]!);
    // Within this session: source × family × every descriptor × content, so
    // a face two claimed families share is registered once.
    const signed = chosen.map((declaration) => {
      const { data, url = '' } = declaration;
      const dataUrl = data === undefined && /^\s*data:/i.test(url);
      return { declaration, signature: faceSignature(state.sourceIndex, declaration), data, url, dataUrl };
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
      const bytes = data !== undefined ? data.byteLength : dataUrl ? dataUrlDecodedByteLength(url) ?? 0 : 0;
      const familyKey = fontFamilyKey(declaration.family);
      const alias = this.aliases.get(familyKey) ?? `ooxml-font-source-${this.serial}-${this.aliases.size}`;
      let face: FontFace;
      let charged = false;
      try {
        // Beyond the realm's retained-byte ceiling the face fails like a
        // face that cannot load (its family warned), nothing else. A
        // prepared face's charge (the same bytes) is taken over: it fit with
        // every earlier face and download charged, so it would fit here too;
        // charging another face first reclaims any prepared face still held.
        const prepared = this.prepared.get(signature);
        this.prepared.delete(signature);
        if (!prepared) chargeRetainedBytes(bytes);
        charged = true;
        if (prepared) {
          face = prepared.face;
          face.family = alias;
        } else {
          face = new FontFace(alias, data ?? cssUrl(url), faceDescriptors(declaration));
        }
        try { set.add(face); }
        catch (error) { set.delete(face); throw error; }
      } catch (error) {
        if (charged) refundRetainedBytes(bytes);
        failures.set(state.name, error instanceof Error ? error.message : String(error));
        continue;
      }
      this.aliases.set(familyKey, alias);
      const entry: OwnedFace = {
        face, declaration, sourceIndex: state.sourceIndex,
        ranges: parseUnicodeRange(declaration.unicodeRange),
        bytes,
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
    warnFailures(fresh);
  }

  private takeSnapshot(): LoadedFontSources {
    const faces: LoadedSourceFace[] = [];
    for (const entry of this.owned.values()) {
      if (entry.face.status !== 'loaded') continue;
      const { declaration } = entry;
      const resourceMetrics = (this.readMetricTables && declaration.data && faceResourceMetrics.get(declaration.data)) || undefined;
      const probeSource = resourceMetrics ? undefined
        : declaration.data ?? (declaration.url === undefined ? undefined : cssUrl(declaration.url));
      faces.push(Object.freeze({
        family: declaration.family,
        weight: declaration.weight ?? 'normal',
        style: declaration.style ?? 'normal',
        stretch: declaration.stretch ?? 'normal',
        unicodeRange: declaration.unicodeRange ?? 'U+0-10FFFF',
        sourceIndex: entry.sourceIndex,
        face: entry.face,
        ...(resourceMetrics ? { resourceMetrics } : {}),
        ...(probeSource === undefined ? {} : { probeSource }),
        ...(probeSource !== undefined && sharesSubsetMetrics(this.sources[entry.sourceIndex], declaration.family)
          ? { sharedSubsetMetrics: true as const } : {}),
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
      families: new Map(this.aliases),
    });
  }
}

const sharesSubsetMetrics = (source: FontSourceDescriptor | undefined, family: string) =>
  source?.kind === 'font-css-api'
  && source.sharedSubsetMetricFamilies?.some((entry) => fontFamilyKey(entry) === fontFamilyKey(family)) === true;

const releasedRunner: FontResolverRunner = () => {
  throw createAbortError('Font source session released');
};
