import {
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  deobfuscateOdttf,
  embeddedFontBytesAreWithinLimit,
  normalizeFontMetricFamily,
  parseOpenTypeResourceMetrics,
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type EmbeddedFontFace,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import {
  parseResourceLimitError,
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
} from '@silurus/ooxml-core/worker';
import type { LayoutDiagnostic } from './layout/types.js';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { sha256 } from '@silurus/ooxml-core/internal/sha256';

// Resource governance for one retained DOCX font set. The byte ceiling matches
// the archive's raw-part cache ceiling; the face ceiling prevents thousands of
// tiny parts from creating an unbounded FontFaceSet. Fetch at most two parts at
// once, then admit them in document order so asynchronous completion cannot
// change which faces fit the budget. These are memory limits, not Office rules.
const MAX_RETAINED_FONT_BYTES = HARD_MAX_RAW_PART_CACHE_BYTES;
const MAX_RETAINED_FONT_FACES = HARD_MAX_RAW_PART_CACHE_ENTRIES;
const MAX_CONCURRENT_FONT_OPERATIONS = 2;

export interface LoadedEmbeddedFonts {
  readonly diagnostics?: readonly LayoutDiagnostic[];
  readonly faces: FontFace[];
  readonly metrics: Readonly<Record<string, ResolvedFontMetric>>;
  /** Successfully loaded authored tuples and their resource-specific Canvas families. */
  readonly routes: readonly LoadedEmbeddedFontRoute[];
}

export interface LoadedEmbeddedFontRoute {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
  readonly resourceIdentity: string;
}

// SHA-256 of the de-obfuscated resource, independent of archive path and load
// order. FontFaceSet is global to a realm: two documents can embed different
// bytes under the same authored name and CSS descriptors. A resource-derived
// private family prevents Canvas from selecting the other document's face.
async function embeddedFamily(data: Uint8Array, authoredFamily: string): Promise<string> {
  // SubtleCrypto is unavailable on ordinary HTTP origins. The synchronous
  // implementation keeps those renderers isolated too, with the same digest.
  let digest: Uint8Array;
  try {
    digest = globalThis.crypto?.subtle
      ? new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256', data.slice().buffer))
      : sha256(data);
  } catch {
    digest = sha256(data);
  }
  const hex = Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('');
  // The authored-name suffix keeps equal bytes under different document names
  // in separate CSS registrations. A name-hash collision is harmless because
  // the bytes, and thus painted geometry, are then identical.
  let nameHash = 0x811c9dc5;
  for (const scalar of authoredFamily.trim().toLowerCase()) {
    nameHash ^= scalar.codePointAt(0)!;
    nameHash = Math.imul(nameHash, 0x01000193);
  }
  return `__ooxml_docx_embedded_${hex}_${(nameHash >>> 0).toString(16)}`;
}

/**
 * Register a document's embedded fonts (ECMA-376 §17.8.3.3-.6) into the active
 * FontFaceSet so the renderer measures and paints text with the authored
 * typeface instead of a substitute.
 *
 * `doc.embeddedFonts` names the obfuscated `.odttf` parts + their `w:fontKey`
 * GUIDs; the bytes are fetched by zip path through `fetchFontBytes` (the docx
 * archive extracts any part, not just images), de-obfuscated per §17.8.1, and
 * added to the set under an isolated family derived from its plaintext bytes.
 * Each `<w:embed*>` style slot becomes one CSS weight/style pair:
 * bold / boldItalic ⇒ `weight: 'bold'`; italic / boldItalic ⇒ `style: 'italic'`.
 *
 * MUST run before pagination (which measures text). No-ops when the document
 * embeds no fonts. Individual part fetches are concurrent; a rejected fetch
 * skips only that face (the rest still register) so one missing part never
 * aborts the whole document. Typed resource-limit failures remain fatal: they
 * enforce archive policy rather than diagnose malformed font content.
 *
 * Returns the shared `FontFace` objects registered for this document (deduped +
 * refcounted in core). The caller ({@link DocxDocument}) holds them and passes
 * them to `unregisterEmbeddedFonts` in `destroy()` so they leave `document.fonts`
 * when the document is discarded, instead of leaking on every open (SPA leak).
 */
export async function loadEmbeddedFonts(
  doc: DocxDocumentModel,
  fetchFontBytes: (partPath: string) => Promise<Uint8Array>,
): Promise<LoadedEmbeddedFonts> {
  const refs = doc.embeddedFonts;
  if (!refs || refs.length === 0) return { faces: [], metrics: {}, routes: [] };

  // A CSS family/weight/style tuple cannot identify two different resources.
  // Preserve document order and admit only the first definition before any
  // fetch or parse, so duplicate declarations cannot amplify resource work or
  // make metrics from one resource describe another face.
  const uniqueRefs: EmbeddedFontRef[] = [];
  const seenRefTuples = new Set<string>();
  for (const ref of refs) {
    const weight = weightForStyle(ref.style) === 'bold' ? 700 : 400;
    const style = styleForStyle(ref.style);
    const tuple = `${normalizeFontMetricFamily(ref.fontName)}:${weight}:${style}`;
    if (seenRefTuples.has(tuple)) continue;
    seenRefTuples.add(tuple);
    uniqueRefs.push(ref);
  }

  type PreparedFace = {
    face: EmbeddedFontFace;
    metric: ResolvedFontMetric | null;
    route: LoadedEmbeddedFontRoute;
  };
  const faces: PreparedFace[] = [];
  let retainedBytes = 0;
  // Bound attempts as well as retained faces: malformed parts must not make a
  // document with thousands of declarations trigger thousands of fetches.
  const admittedRefCount = Math.min(uniqueRefs.length, MAX_RETAINED_FONT_FACES);
  let resourceSkipped = uniqueRefs.length - admittedRefCount;
  for (let start = 0; start < admittedRefCount;
    start += MAX_CONCURRENT_FONT_OPERATIONS) {
    const batch = uniqueRefs.slice(start, Math.min(start + MAX_CONCURRENT_FONT_OPERATIONS, admittedRefCount));
    const fetched = await Promise.all(batch.map(async (ref) => {
      try { return await fetchFontBytes(ref.partPath); }
      catch (error) {
        if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
        // Worker rendering reads WASM directly, before bridge deserialization.
        const resourceError = parseResourceLimitError(error);
        if (resourceError) throw resourceError;
        return null;
      }
    }));
    for (let index = 0; index < batch.length; index++) {
      const ref = batch[index];
      const bytes = fetched[index];
      if (!bytes) continue;
      if (!embeddedFontBytesAreWithinLimit(bytes)
        || bytes.byteLength > MAX_RETAINED_FONT_BYTES - retainedBytes) {
        resourceSkipped++;
        continue;
      }
      try {
        const odttf = ref.partPath.toLowerCase().endsWith('.odttf');
        // Route identity must exist before registration. If de-obfuscation or
        // digesting fails, never fall back to the shared authored CSS name.
        const data = odttf ? deobfuscateOdttf(bytes, ref.fontKey ?? '') : bytes;
        const family = await embeddedFamily(data, ref.fontName);
        // FontFace registration and metric extraction use the same plaintext.
        const face = {
          family,
          bytes: data,
          odttf: false,
          fontKey: '',
          weight: weightForStyle(ref.style),
          style: styleForStyle(ref.style),
        } satisfies EmbeddedFontFace;
        let metric: ResolvedFontMetric | null = null;
        try {
          const openType = parseOpenTypeResourceMetrics(data);
          // The wordOpenTypeAutoLineRatios projection requires this face's
          // OS/2 code-page class and usWin data; cmap coverage is a separate
          // question. A face without declared code-page ranges does not gain
          // this authority merely because it contains a CJK glyph.
          const line = openType?.farEastCodePage == null
            ? null
            : wordOpenTypeAutoLineRatios({
                ...openType,
                farEastCodePage: openType.farEastCodePage,
              });
          if (line || openType?.averageCharWidthRatio != null) {
            metric = Object.freeze({
              family,
              requestedFamily: ref.fontName,
              weight: face.weight === 'bold' ? 700 : 400,
              style: face.style,
              sourceIdentity: `embedded:${family}`,
              synthesized: false,
              ...(line ? {
                lineHeightRatio: line.lineHeightRatio,
                designAscentRatio: line.designAscentRatio,
                designDescentRatio: line.designDescentRatio,
              } : {}),
              ...(openType?.averageCharWidthRatio == null ? {} : {
                averageCharWidthRatio: openType.averageCharWidthRatio,
              }),
              unicodeRanges: openType?.unicodeRanges ?? [],
              ...(line && openType?.hasEastAsianCmap
                ? { eastAsianLineHeightRatio: line.lineHeightRatio }
                : {}),
            });
          }
        } catch {
          // Registration owns the malformed-font diagnostic. Metrics are an
          // optional derivative and must never make a loadable face fatal.
        }
        faces.push({
          face, metric,
          route: {
            requestedFamily: ref.fontName,
            resolvedFamily: face.family,
            weight: face.weight === 'bold' ? 700 : 400,
            style: face.style,
            resourceIdentity: `embedded:${face.family}`,
          },
        });
        retainedBytes += bytes.byteLength;
      } catch {
        // A missing / unreadable part: skip this face, keep the rest.
        continue;
      }
    }
  }

  // These are library resource ceilings, not an Office font-selection rule.
  // Retain a diagnostic even when no face fits; callers render with fallback.
  const diagnostics: readonly LayoutDiagnostic[] = resourceSkipped > 0
    ? Object.freeze([Object.freeze({
        code: 'MISSING_RESOURCE' as const,
        severity: 'warning' as const,
        message: `Skipped ${resourceSkipped} embedded font face(s) under the document font resource limits; fallback fonts are used`,
      })])
    : Object.freeze([]);

  const loadable = faces;
  if (loadable.length === 0) return { faces: [], metrics: {}, routes: [], diagnostics };
  // FontFace.load() runs concurrently within one core registration call. Batch
  // registration as well as fetching so browser font decoding respects the
  // same per-document concurrency limit. Release earlier batches if a later
  // registration unexpectedly fails before the caller can own them.
  const loadedFaces: FontFace[] = [];
  const heldFaces = new Set<FontFace>();
  try {
    for (let start = 0; start < loadable.length; start += MAX_CONCURRENT_FONT_OPERATIONS) {
      const batch = loadable.slice(start, start + MAX_CONCURRENT_FONT_OPERATIONS);
      for (const face of await registerEmbeddedFonts(batch.map((entry) => entry.face))) {
        if (heldFaces.has(face)) unregisterEmbeddedFonts([face]);
        else { heldFaces.add(face); loadedFaces.push(face); }
      }
    }
  } catch (error) {
    unregisterEmbeddedFonts(loadedFaces);
    throw error;
  }
  const loadedTuples = new Set(loadedFaces.map((face) => {
    const family = face.family.trim().replace(/^(['"])(.*)\1$/, '$2');
    const weight = face.weight === 'bold' ? 700 : Number(face.weight) || 400;
    const style = face.style === 'italic' ? 'italic' : 'normal';
    return `${normalizeFontMetricFamily(family)}:${weight}:${style}`;
  }));
  // Null prototype: keys are document font names (e.g. `__proto__`).
  const metrics = Object.create(null) as Record<string, ResolvedFontMetric>;
  const routes: LoadedEmbeddedFontRoute[] = [];
  for (const entry of loadable) {
    const family = normalizeFontMetricFamily(entry.face.family);
    const weight = entry.face.weight === 'bold' ? 700 : 400;
    const style = entry.face.style;
    const tuple = `${family}:${weight}:${style}`;
    if (!loadedTuples.has(tuple)) continue;
    routes.push(Object.freeze(entry.route));
    if (!entry.metric) continue;
    const requested = normalizeFontMetricFamily(entry.route.requestedFamily);
    metrics[weight === 400 && style === 'normal' ? requested : `${requested}:${weight}:${style}`] = entry.metric;
  }
  return { faces: loadedFaces, metrics: Object.freeze(metrics), routes: Object.freeze(routes), diagnostics };
}

/** bold / boldItalic slots ⇒ CSS `font-weight: bold`; otherwise `normal`. */
function weightForStyle(style: EmbeddedFontRef['style']): 'normal' | 'bold' {
  return style === 'bold' || style === 'boldItalic' ? 'bold' : 'normal';
}

/** italic / boldItalic slots ⇒ CSS `font-style: italic`; otherwise `normal`. */
function styleForStyle(style: EmbeddedFontRef['style']): 'normal' | 'italic' {
  return style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal';
}
