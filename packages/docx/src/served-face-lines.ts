import {
  loadNativeRouteSentinels,
  measureFontFaceNormalLine,
  unloadOfficeFontFallbacks,
  type LocalFontNormalLine,
} from '@silurus/ooxml-core';
import {
  fontFamilyKey,
  fontSourceFaceIdentity,
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
 * from the same source for the ownership judgement.
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

export function servedFaceLineRequests(loaded: LoadedFontSources): ServedFaceLineRequest[] {
  return loaded.faces.flatMap((face, index) => face.resourceMetrics || face.probeSource === undefined ? [] : [{
    index,
    source: face.probeSource,
    weight: face.weight,
    style: face.style,
    stretch: face.stretch,
  }]);
}

let probeSerial = 0;

function registerProbe(request: ServedFaceLineRequest, fontSet: FontFaceSet): FontFace {
  const family = `__ooxml_served_face_${++probeSerial}`;
  return retainFace(`served-face:${family}`, fontSet, () => {
    const created = new FontFace(family, request.source, {
      weight: request.weight, style: request.style, stretch: request.stretch,
    });
    fontSet.add(created);
    return created;
  }).face;
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

function measureTuples(face: FontFace, owner: Document): Record<string, LocalFontNormalLine | null> {
  const context = owner.createElement('canvas').getContext('2d');
  const lines: Record<string, LocalFontNormalLine | null> = Object.create(null);
  for (const { weight, style } of WML_TUPLES) {
    lines[tupleKey(weight, style)] = context
      ? measureFontFaceNormalLine(face, weight, style, owner.fonts, context, PROBE_EM_PX, owner)
      : null;
  }
  return lines;
}

/** DOM realm, worker mode: measure each request with a transient private
 * registration (released before returning). */
export async function measureServedFaceLines(
  requests: readonly ServedFaceLineRequest[],
  signal?: AbortSignal,
): Promise<ServedFaceLineMeasurement[]> {
  const owner = typeof document === 'undefined' ? null : document;
  if (!owner?.body) return requests.map(({ index }) => ({ index, lines: {} }));
  return Promise.all(requests.map(async (request) => {
    const face = registerProbe(request, owner.fonts);
    try {
      return { index: request.index, lines: await loaded(face, signal) ? measureTuples(face, owner) : {} };
    } finally {
      releaseFaces([face]);
    }
  }));
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
  measured?: readonly ServedFaceLineMeasurement[],
  signal?: AbortSignal,
): Promise<PreparedServedFaceLines & { faces: FontFace[] }> {
  const requests = sources ? servedFaceLineRequests(sources) : [];
  const owner = typeof document !== 'undefined' && document.body && document.fonts === fontSet ? document : null;
  // Application faces are measured only where they live: the DOM realm.
  const appFaces = owner ? appRegisteredFaces(sources, fontSet, families) : [];
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
  await Promise.all(requests.map(async (request) => {
    const face = registerProbe(request, fontSet);
    if (!await loaded(face, signal)) {
      releaseFaces([face]);
      return;
    }
    faces.push(face);
    const measurement = owner ? measureTuples(face, owner)
      : measured?.find((candidate) => candidate.index === request.index)?.lines;
    if (measurement) faceLines.set(request.index, { ownershipFamily: face.family, lines: measurement });
  }));
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
  for (const face of appFaces) {
    familyLines.set(fontFamilyKey(unquoted(face.family)), { ownershipFamily: unquoted(face.family), lines: measureTuples(face, owner!) });
  }
  if (faceLines.size === 0 && familyLines.size === 0) {
    unloadOfficeFontFallbacks(faces);
    return { ...NO_SERVED_FACE_LINES, faces: [] };
  }
  return { faceLines, familyLines, faces };
}

const unquoted = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2');

/** Loaded application FontFaces of `families` that are their family's only
 * loaded face, cover U+0020, and are not this session's own faces. */
export function appRegisteredFaces(
  sources: LoadedFontSources | undefined,
  fontSet: FontFaceSet | null,
  families: readonly string[],
): FontFace[] {
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return [];
  const wanted = new Set(families.map(fontFamilyKey));
  const session = new Set(sources?.faces.map((face) => face.face) ?? []);
  const byFamily = new Map<string, FontFace[]>();
  for (const face of fontSet) {
    if (face.status !== 'loaded') continue;
    const key = fontFamilyKey(unquoted(face.family));
    if (wanted.has(key)) byFamily.set(key, [...byFamily.get(key) ?? [], face]);
  }
  const single: FontFace[] = [];
  for (const faces of byFamily.values()) {
    const [face] = faces;
    // A face any font-source session registered is a served face, not one
    // the application registered.
    if (faces.length !== 1 || !face || session.has(face) || fontSourceFaceIdentity(face) !== undefined) continue;
    const ranges = parseUnicodeRange(face.unicodeRange) ?? [[0, 0x10ffff]];
    if (ranges.some(([lo, hi]) => lo <= 0x20 && 0x20 <= hi)) single.push(face);
  }
  return single;
}
