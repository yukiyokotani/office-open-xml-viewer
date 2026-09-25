import {
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
  HARD_MAX_RAW_PART_CACHE_BYTES,
  HARD_MAX_RAW_PART_CACHE_ENTRIES,
} from '@silurus/ooxml-core/worker';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';

// Resource governance for one retained DOCX font set. The byte ceiling matches
// the archive's raw-part cache ceiling; the face ceiling prevents thousands of
// tiny parts from creating an unbounded FontFaceSet. Fetch at most two parts at
// once, then admit them in document order so asynchronous completion cannot
// change which faces fit the budget. These are memory limits, not Office rules.
const MAX_RETAINED_FONT_BYTES = HARD_MAX_RAW_PART_CACHE_BYTES;
const MAX_RETAINED_FONT_FACES = HARD_MAX_RAW_PART_CACHE_ENTRIES;
const MAX_CONCURRENT_FONT_OPERATIONS = 2;

export interface LoadedEmbeddedFonts {
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
  for (const scalar of authoredFamily.trim().toLocaleLowerCase('en-US')) {
    nameHash ^= scalar.codePointAt(0)!;
    nameHash = Math.imul(nameHash, 0x01000193);
  }
  return `__ooxml_docx_embedded_${hex}_${(nameHash >>> 0).toString(16)}`;
}

const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

function sha256(bytes: Uint8Array): Uint8Array {
  const h = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);
  const words = new Uint32Array(64);
  const total = Math.ceil((bytes.length + 9) / 64) * 64;
  const bitLength = bytes.length * 8;
  const paddedByte = (index: number): number => index < bytes.length ? bytes[index]
    : index === bytes.length ? 0x80
      : index >= total - 8 ? Math.floor(bitLength / 2 ** ((total - index - 1) * 8)) & 0xff : 0;
  const rotate = (value: number, bits: number) => (value >>> bits) | (value << (32 - bits));
  for (let offset = 0; offset < total; offset += 64) {
    for (let i = 0; i < 16; i++) {
      const start = offset + i * 4;
      words[i] = (paddedByte(start) << 24) | (paddedByte(start + 1) << 16)
        | (paddedByte(start + 2) << 8) | paddedByte(start + 3);
    }
    for (let i = 16; i < 64; i++) {
      const p = words[i - 15];
      const q = words[i - 2];
      const s0 = rotate(p, 7) ^ rotate(p, 18) ^ (p >>> 3);
      const s1 = rotate(q, 17) ^ rotate(q, 19) ^ (q >>> 10);
      words[i] = (words[i - 16] + s0 + words[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, j] = h;
    for (let i = 0; i < 64; i++) {
      const s1 = rotate(e, 6) ^ rotate(e, 11) ^ rotate(e, 25);
      const choose = (e & f) ^ (~e & g);
      const t1 = (j + s1 + choose + SHA256_K[i] + words[i]) >>> 0;
      const s0 = rotate(a, 2) ^ rotate(a, 13) ^ rotate(a, 22);
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (s0 + majority) >>> 0;
      j = g; g = f; f = e; e = (d + t1) >>> 0;
      d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    h[0] = (h[0] + a) >>> 0; h[1] = (h[1] + b) >>> 0;
    h[2] = (h[2] + c) >>> 0; h[3] = (h[3] + d) >>> 0;
    h[4] = (h[4] + e) >>> 0; h[5] = (h[5] + f) >>> 0;
    h[6] = (h[6] + g) >>> 0; h[7] = (h[7] + j) >>> 0;
  }
  const result = new Uint8Array(32);
  const view = new DataView(result.buffer);
  h.forEach((value, index) => view.setUint32(index * 4, value));
  return result;
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
 * aborts the whole document.
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
      catch { return null; }
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
          // Word for Mac's observed auto-line class follows OS/2 code-page
          // bits 17–20, including for Latin glyphs; cmap coverage is a separate
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

  if (resourceSkipped > 0) {
    console.warn(`[ooxml] skipped ${resourceSkipped} embedded font face(s) under the document font resource limits`);
  }

  const loadable = faces;
  if (loadable.length === 0) return { faces: [], metrics: {}, routes: [] };
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
  const metrics: Record<string, ResolvedFontMetric> = {};
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
  return { faces: loadedFaces, metrics: Object.freeze(metrics), routes: Object.freeze(routes) };
}

/** bold / boldItalic slots ⇒ CSS `font-weight: bold`; otherwise `normal`. */
function weightForStyle(style: EmbeddedFontRef['style']): 'normal' | 'bold' {
  return style === 'bold' || style === 'boldItalic' ? 'bold' : 'normal';
}

/** italic / boldItalic slots ⇒ CSS `font-style: italic`; otherwise `normal`. */
function styleForStyle(style: EmbeddedFontRef['style']): 'normal' | 'italic' {
  return style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal';
}
