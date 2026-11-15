import {
  loadNativeRouteSentinels,
  measureFontFaceNormalLines,
  unloadOfficeFontFallbacks,
  type LocalFontNormalLine,
} from '@silurus/ooxml-core';
import {
  fontFamilyKey,
  parseUnicodeRange,
  releaseFaces,
  retainFace,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';

/**
 * Measured normal lines of served faces whose tables cannot be read
 * (source-font-metrics.ts uses tables whenever it has them).
 *
 * A WOFF2 file where the engine cannot inflate Brotli, or an inline data: URL
 * face, is registered once more, alone, under a private family
 * (`probeSource`: the same bytes or the same data: URL, so no request). With one face in its family, both the CSS
 * `line-height: normal` strut and the Canvas font box are that face's own, so
 * the measurement (core local-line-metrics.ts) belongs to exactly this face,
 * whatever other subsets its served family has. The private family stays
 * registered in the layout realm: route sentinels judge, per cluster, that it
 * draws the glyphs (native-route-line-metrics.ts), and the declared
 * `unicode-range` limits the metric to the scalars the served family routes
 * to this face.
 *
 * In worker mode the DOM realm measures (a worker has no layout strut) and
 * the worker receives plain numbers; it registers its own private family
 * from the same source for the ownership judgement, while the DOM realm
 * measures.
 *
 * Subsets of one style (`LoadedSourceFace.sharedSubsetMetrics`). A source
 * may guarantee that its `unicode-range` subsets of one served family with
 * equal descriptors keep one font file's vertical metric tables (core
 * FontCssApiSource.sharedSubsetMetricFamilies: the Google Fonts CSS2 API,
 * observed for every subset of its catalog). This applies only when every
 * loaded face of the served family comes from that one source, since the
 * session's private family holds the faces of every source. Their measurements are then equal, so
 * the style is measured once, through the subset whose range holds U+0020
 * (the style's first available font, CSS Fonts 4 §5: the face whose metrics
 * CSS gives the family's normal line), and every loaded subset of the style
 * takes that measurement. A style none of whose loaded subsets holds U+0020
 * is measured per face, as above. Ownership of such a style's clusters is
 * judged on the session's own private family of the served family (the
 * family that paints): a cluster is owned when the face CSS selects for it
 * among the style's subsets draws it, which, for equal geometry and the
 * declared `unicode-range` gate, gives the per-subset judgement's answer, so
 * no subset is registered again to be judged and the measuring probe is
 * released at once.
 *
 * Cost (library resource policy): a probe is a pure function of its source
 * bytes and descriptors in one realm, so one registration (one browser
 * decode) serves every document of the realm holding the same downloaded
 * buffer; every probe of a document is measured in one batch (core measureFontFaceNormalLines), a few layouts in all rather
 * than several per face.
 */

/** The longest a probe registration may take to load: the font-source
 * face-loading ceiling. A probe is optional; one that does not load in time
 * is released and its face keeps the Canvas box. */
export const SERVED_FACE_PROBE_TIMEOUT_MS = 15_000;

/** Measured at a power-of-two em; local-line-metrics checks linear scaling. */
const PROBE_EM_PX = 1024;

/** WordprocessingML tuples: regular/bold × upright/italic. */
export const WML_TUPLES = [
  { weight: 400, style: 'normal' }, { weight: 700, style: 'normal' },
  { weight: 400, style: 'italic' }, { weight: 700, style: 'italic' },
] as const;

export const tupleKey = (weight: number, style: string) => `${weight}:${style}`;

/** One face to measure, identified by its index in the snapshot. */
export interface ServedFaceLineRequest {
  readonly index: number;
  readonly source: string | ArrayBuffer;
  readonly weight: string;
  readonly style: string;
  readonly stretch: string;
  /** Snapshot indices of the subsets of a shared-metric style this face
   * measures for (itself included); absent for a face measured for itself. */
  readonly members?: readonly number[];
}

export interface ServedFaceLineMeasurement {
  readonly index: number;
  /** By {@link tupleKey}; null when unmeasurable. */
  readonly lines: Readonly<Record<string, LocalFontNormalLine | null>>;
}

/** Normal lines of one served face, ready for metric projection. */
export interface ServedFaceLine {
  /** The private family whose clusters the route sentinels judge. */
  readonly ownershipFamily: string;
  readonly lines: Readonly<Record<string, LocalFontNormalLine | null>>;
}

const holdsSpace = (unicodeRange: string) =>
  (parseUnicodeRange(unicodeRange) ?? [[0, 0x10ffff] as const]).some(([lo, hi]) => lo <= 0x20 && 0x20 <= hi);

/** The faces to measure: each face without tables, except that the loaded
 * subsets of a shared-metric style are measured once, through the last
 * registered of them whose range holds U+0020 (see the module comment). */
export function servedFaceLineRequests(loaded: LoadedFontSources): ServedFaceLineRequest[] {
  const requests: ServedFaceLineRequest[] = [];
  const styles = new Map<string, number[]>();
  const request = (index: number, members?: readonly number[]): ServedFaceLineRequest => {
    const face = loaded.faces[index]!;
    return {
      index, source: face.probeSource!, weight: face.weight, style: face.style, stretch: face.stretch,
      ...(members ? { members } : {}),
    };
  };
  // The painting family judges a style only when every loaded face of the
  // served family is a shared-metric subset of one source: the session's
  // private family holds every source's faces of that served family.
  const sourceOf = new Map<string, number | null>();
  for (const face of loaded.faces) {
    const key = fontFamilyKey(face.family);
    const source = face.sharedSubsetMetrics ? face.sourceIndex : null;
    sourceOf.set(key, sourceOf.has(key) && sourceOf.get(key) !== source ? null : source);
  }
  loaded.faces.forEach((face, index) => {
    if (face.resourceMetrics || face.probeSource === undefined) return;
    if (!face.sharedSubsetMetrics || sourceOf.get(fontFamilyKey(face.family)) == null) {
      requests.push(request(index));
      return;
    }
    const key = [face.sourceIndex, fontFamilyKey(face.family), face.weight, face.style, face.stretch].join('\u0000');
    const members = styles.get(key);
    if (members) members.push(index);
    else styles.set(key, [index]);
  });
  for (const members of styles.values()) {
    const first = [...members].reverse().find((index) => holdsSpace(loaded.faces[index]!.unicodeRange));
    if (first === undefined) requests.push(...members.map((index) => request(index)));
    else requests.push(request(first, Object.freeze(members)));
  }
  return requests.sort((a, b) => a.index - b.index);
}

let probeSerial = 0;

/** Serial of each probe source buffer in this realm (see {@link registerProbe}). */
const probeBuffers = new WeakMap<ArrayBuffer, number>();
let probeBufferSerial = 0;

/** The registration key of a probe: its source and descriptors. A probe is a
 * pure function of its source bytes and descriptors in one realm (its own
 * family holds it alone), so equal keys measure and own exactly alike.
 * Downloaded bytes are one buffer for every session of the realm
 * (core face-download.ts), so documents sharing a face share its probe. An
 * inline data: URL is not shared (no cheap identity). */
function probeKey(request: ServedFaceLineRequest): string | undefined {
  const { source } = request;
  if (typeof source === 'string') return undefined;
  let serial = probeBuffers.get(source);
  if (serial === undefined) { serial = ++probeBufferSerial; probeBuffers.set(source, serial); }
  return `served-face:bytes:${serial}\u0000${request.weight}\u0000${request.style}\u0000${request.stretch}`;
}

/** Register (or join) the probe of `request` in `fontSet`. Each call holds
 * one reference, released with releaseFaces. */
function registerProbe(request: ServedFaceLineRequest, fontSet: FontFaceSet): FontFace {
  const key = probeKey(request);
  return retainFace(key ?? `served-face:${++probeSerial}`, fontSet, () => {
    const created = new FontFace(`__ooxml_served_face_${++probeSerial}`, request.source, {
      weight: request.weight, style: request.style, stretch: request.stretch,
    });
    fontSet.add(created);
    return created;
  }).face;
}

/** One probe per distinct registration key of `requests`, each held once:
 * the probe of every request, by position. */
function registerProbes(requests: readonly ServedFaceLineRequest[], fontSet: FontFaceSet): FontFace[] {
  const byKey = new Map<string, FontFace>();
  return requests.map((request) => {
    const key = probeKey(request);
    const held = key === undefined ? undefined : byKey.get(key);
    if (held) return held;
    const face = registerProbe(request, fontSet);
    if (key !== undefined) byKey.set(key, face);
    return face;
  });
}

/** Whether `face` loads within {@link SERVED_FACE_PROBE_TIMEOUT_MS} and
 * before `signal` aborts. `FontFace.load()` cannot be cancelled: the wait
 * stops, and the caller releases the registration. */
async function loaded(face: FontFace, signal?: AbortSignal): Promise<boolean> {
  if (signal?.aborted) return false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  const stop = new Promise<false>((resolve) => {
    timer = setTimeout(() => resolve(false), SERVED_FACE_PROBE_TIMEOUT_MS);
    onAbort = () => resolve(false);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
  try {
    return await Promise.race([
      face.load().then(() => face.status === 'loaded', () => false),
      stop,
    ]);
  } finally {
    clearTimeout(timer);
    if (onAbort) signal?.removeEventListener('abort', onAbort);
  }
}

type TupleLines = Readonly<Record<string, LocalFontNormalLine | null>>;

/** Every WordprocessingML tuple of each face, measured in one batch (core
 * measureFontFaceNormalLines: at most a few layouts for the whole list).
 * Every document measures its own probes, even a probe it shares: a
 * measurement describes the page as it is now (a hidden body measures
 * nothing), so another document's earlier answer never stands in for it. */
function measureTuples(faces: readonly FontFace[], owner: Document): TupleLines[] {
  const distinct = [...new Set(faces)];
  const context = distinct.length ? owner.createElement('canvas').getContext('2d') : null;
  const measured = context
    ? measureFontFaceNormalLines(distinct.flatMap((face) => WML_TUPLES.map(({ weight, style }) => ({ face, weight, style }))),
      owner.fonts, context, PROBE_EM_PX, owner)
    : [];
  const byFace = new Map<FontFace, TupleLines>();
  distinct.forEach((face, index) => {
    const lines: Record<string, LocalFontNormalLine | null> = Object.create(null);
    WML_TUPLES.forEach(({ weight, style }, tuple) => {
      lines[tupleKey(weight, style)] = measured[index * WML_TUPLES.length + tuple] ?? null;
    });
    byFace.set(face, lines);
  });
  return faces.map((face) => byFace.get(face)!);
}

/** DOM realm, worker mode: measure each request with a transient private
 * registration (released before returning). */
export async function measureServedFaceLines(
  requests: readonly ServedFaceLineRequest[],
  signal?: AbortSignal,
): Promise<ServedFaceLineMeasurement[]> {
  const owner = typeof document === 'undefined' ? null : document;
  if (!owner?.body) return requests.map(({ index }) => ({ index, lines: {} }));
  const faces = registerProbes(requests, owner.fonts);
  try {
    const ready = await Promise.all(faces.map((face) => loaded(face, signal)));
    const measured = measureTuples(faces.filter((_, position) => ready[position]), owner);
    let next = 0;
    return requests.map((request, position) => ({ index: request.index, lines: ready[position] ? measured[next++]! : {} }));
  } finally {
    // After the caller has posted the answer: removing hundreds of faces from
    // the set is not on the worker's path.
    setTimeout(() => releaseFaces(new Set(faces)), 0);
  }
}

/** Measured lines for one document's layout realm. */
export interface PreparedServedFaceLines {
  /** By snapshot face index. */
  readonly faceLines: ReadonlyMap<number, ServedFaceLine>;
  /** Application-registered faces, by family key (fontFamilyKey). */
  readonly familyLines: ReadonlyMap<string, ServedFaceLine>;
}

export const NO_SERVED_FACE_LINES: PreparedServedFaceLines = Object.freeze({
  faceLines: new Map(), familyLines: new Map(),
});

/**
 * Layout realm: register each face's private family (kept until the document
 * is destroyed, released through `faces`), load the route sentinels, and pair
 * the private families with measurements: measured here when this realm has
 * the DOM (main mode), else the DOM realm's `measured` answer (worker mode).
 * `families` limits application-registered faces to the document's families.
 * Each probe load is bounded ({@link SERVED_FACE_PROBE_TIMEOUT_MS}); a probe
 * that does not load is released at once, and once `signal` aborts every
 * registration made here is released and the wait ends with no lines.
 */
export async function prepareServedFaceLines(
  sources: LoadedFontSources | undefined,
  fontSet: FontFaceSet | null,
  families: readonly string[],
  measured?: readonly ServedFaceLineMeasurement[] | PromiseLike<readonly ServedFaceLineMeasurement[]>,
  signal?: AbortSignal,
): Promise<PreparedServedFaceLines & { faces: FontFace[] }> {
  const requests = sources ? servedFaceLineRequests(sources) : [];
  const owner = typeof document !== 'undefined' && document.body && document.fonts === fontSet ? document : null;
  // Application faces are measured only where they live: the DOM realm.
  const appFaces = owner ? appRegisteredFaces(fontSet, families) : [];
  if (!fontSet || (requests.length === 0 && appFaces.length === 0)) return { ...NO_SERVED_FACE_LINES, faces: [] };
  if (signal?.aborted) return { ...NO_SERVED_FACE_LINES, faces: [] };
  const sentinels = await loadNativeRouteSentinels(fontSet);
  if (!sentinels) return { ...NO_SERVED_FACE_LINES, faces: [] };
  const faces: FontFace[] = [...sentinels];
  if (signal?.aborted) {
    unloadOfficeFontFallbacks(faces);
    return { ...NO_SERVED_FACE_LINES, faces: [] };
  }
  const faceLines = new Map<number, ServedFaceLine>();
  const familyLines = new Map<string, ServedFaceLine>();
  // Probes: every face measured here (main mode), and every face judged on
  // its own probe (not a shared-metric style). All are registered before any
  // load is awaited; the loaded ones are measured together in one batch.
  const probed = requests.filter((request) => owner || !request.members);
  const probes = registerProbes(probed, fontSet);
  const ready = await Promise.all(probes.map((face) => loaded(face, signal)));
  const probeOf = new Map<ServedFaceLineRequest, FontFace>();
  probed.forEach((request, position) => { if (ready[position]) probeOf.set(request, probes[position]!); });
  // Kept for the document: the loaded probes of faces judged on their own.
  const kept = new Set(probed.flatMap((request) => !request.members && probeOf.has(request) ? [probeOf.get(request)!] : []));
  faces.push(...kept);
  const measurable = requests.filter((request) => probeOf.has(request) || (!owner && request.members));
  // Worker mode: the DOM realm measures while this realm registers and
  // loads its own probes.
  const answered = owner || signal?.aborted ? undefined : await measured;
  const byIndex = new Map(answered?.map((measurement) => [measurement.index, measurement.lines]));
  const measurements: (TupleLines | undefined)[] = owner && !signal?.aborted
    ? measureTuples(measurable.map((request) => probeOf.get(request)!), owner)
    : measurable.map((request) => byIndex.get(request.index));
  // A shared-metric style's measuring probe is not needed once measured.
  releaseFaces(new Set(probes.filter((face) => !kept.has(face))));
  measurable.forEach((request, position) => {
    const lines = measurements[position];
    if (!lines) return;
    if (!request.members) {
      faceLines.set(request.index, { ownershipFamily: probeOf.get(request)!.family, lines });
      return;
    }
    // The session's private family of the served family paints the style.
    const alias = sources?.families.get(fontFamilyKey(sources.faces[request.index]!.family));
    if (alias) for (const member of request.members) faceLines.set(member, { ownershipFamily: alias, lines });
  });
  if (signal?.aborted) {
    unloadOfficeFontFallbacks(faces);
    return { ...NO_SERVED_FACE_LINES, faces: [] };
  }
  // A loaded FontFace the application registered for an authored family
  // (`css` inventory). Its source is not readable, so it cannot be registered
  // alone; it is measured under its own family only when it is that family's
  // single loaded face and its unicode-range holds U+0020, so it is the
  // family's first available font (CSS Fonts 4 §5) and both the strut and the
  // Canvas box are its own.
  const appLines = appFaces.length ? measureTuples(appFaces, owner!) : [];
  appFaces.forEach((face, position) => {
    familyLines.set(fontFamilyKey(unquoted(face.family)), { ownershipFamily: unquoted(face.family), lines: appLines[position]! });
  });
  if (faceLines.size === 0 && familyLines.size === 0) {
    unloadOfficeFontFallbacks(faces);
    return { ...NO_SERVED_FACE_LINES, faces: [] };
  }
  return { faceLines, familyLines, faces };
}

const unquoted = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2');

/** Loaded application FontFaces of `families` that are their family's only
 * loaded face and cover U+0020. Font sources register under private families
 * (core source-session.ts), so a face declaring a document family is always
 * the application's own. */
export function appRegisteredFaces(
  fontSet: FontFaceSet | null,
  families: readonly string[],
): FontFace[] {
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return [];
  const wanted = new Set(families.map(fontFamilyKey));
  const byFamily = new Map<string, FontFace[]>();
  for (const face of fontSet) {
    if (face.status !== 'loaded') continue;
    const key = fontFamilyKey(unquoted(face.family));
    if (wanted.has(key)) byFamily.set(key, [...byFamily.get(key) ?? [], face]);
  }
  const single: FontFace[] = [];
  for (const faces of byFamily.values()) {
    const [face] = faces;
    if (faces.length !== 1 || !face) continue;
    const ranges = parseUnicodeRange(face.unicodeRange) ?? [[0, 0x10ffff]];
    if (ranges.some(([lo, hi]) => lo <= 0x20 && 0x20 <= hi)) single.push(face);
  }
  return single;
}
