import {
  deobfuscateOdttf,
  embeddedFontBytesAreWithinLimit,
  normalizeFontMetricFamily,
  parseOpenTypeLineMetrics,
  registerEmbeddedFonts,
  type EmbeddedFontFace,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel, EmbeddedFontRef } from './types';
import { wordOpenTypeEastAsianSingleLineRatio } from './layout/line-compatibility.js';

export interface LoadedEmbeddedFonts {
  readonly faces: FontFace[];
  readonly metrics: Readonly<Record<string, ResolvedFontMetric>>;
}

/**
 * Register a document's embedded fonts (ECMA-376 §17.8.3.3-.6) into the active
 * FontFaceSet so the renderer measures and paints text with the authored
 * typeface instead of a substitute.
 *
 * `doc.embeddedFonts` names the obfuscated `.odttf` parts + their `w:fontKey`
 * GUIDs; the bytes are fetched by zip path through `fetchFontBytes` (the docx
 * archive extracts any part, not just images), de-obfuscated per §17.8.1, and
 * added to the set under the exact document
 * font name. Each `<w:embed*>` style slot becomes one CSS weight/style pair:
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
  if (!refs || refs.length === 0) return { faces: [], metrics: {} };

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

  const faces = await Promise.all(
    uniqueRefs.map(async (ref): Promise<{
      face: EmbeddedFontFace;
      metric: ResolvedFontMetric | null;
    } | null> => {
      try {
        const bytes = await fetchFontBytes(ref.partPath);
        const odttf = ref.partPath.toLowerCase().endsWith('.odttf');
        const sourceFace = {
          family: ref.fontName,
          bytes,
          odttf,
          fontKey: ref.fontKey,
          weight: weightForStyle(ref.style),
          style: styleForStyle(ref.style),
        } satisfies EmbeddedFontFace;
        // Keep resource governance ahead of both de-obfuscation and OpenType
        // parsing. Core will reject the same face and emit its normal warning.
        if (!embeddedFontBytesAreWithinLimit(bytes)) {
          return { face: sourceFace, metric: null };
        }

        let face = sourceFace;
        let metric: ResolvedFontMetric | null = null;
        try {
          const data = odttf ? deobfuscateOdttf(bytes, ref.fontKey ?? '') : bytes;
          // Pass plaintext to core so a valid .odttf part is expanded exactly
          // once. FontFace registration and metric extraction now share the
          // same resource bytes.
          face = { ...sourceFace, bytes: data, odttf: false, fontKey: '' };
          const openType = parseOpenTypeLineMetrics(data);
          const eastAsianLineHeightRatio = openType?.hasEastAsianCmap
            ? wordOpenTypeEastAsianSingleLineRatio(openType)
            : 0;
          if (eastAsianLineHeightRatio > 0) {
            metric = Object.freeze({
              family: ref.fontName,
              requestedFamily: ref.fontName,
              weight: face.weight === 'bold' ? 700 : 400,
              style: face.style,
              sourceIdentity: `embedded:${ref.partPath}`,
              synthesized: false,
              eastAsianLineHeightRatio,
            });
          }
        } catch {
          // Registration owns the malformed-font diagnostic. Metrics are an
          // optional derivative and must never make a loadable face fatal.
        }
        return { face, metric };
      } catch {
        // A missing / unreadable part: skip this face, keep the rest.
        return null;
      }
    }),
  );

  const loadable = faces.filter((entry): entry is NonNullable<typeof entry> => entry !== null);
  if (loadable.length === 0) return { faces: [], metrics: {} };
  const loadedFaces = await registerEmbeddedFonts(loadable.map((entry) => entry.face));
  const loadedTuples = new Set(loadedFaces.map((face) => {
    const family = face.family.trim().replace(/^(['"])(.*)\1$/, '$2');
    const weight = face.weight === 'bold' ? 700 : Number(face.weight) || 400;
    const style = face.style === 'italic' ? 'italic' : 'normal';
    return `${normalizeFontMetricFamily(family)}:${weight}:${style}`;
  }));
  const metrics: Record<string, ResolvedFontMetric> = {};
  for (const entry of loadable) {
    if (!entry.metric) continue;
    const family = normalizeFontMetricFamily(entry.face.family);
    const weight = entry.face.weight === 'bold' ? 700 : 400;
    const style = entry.face.style;
    const tuple = `${family}:${weight}:${style}`;
    if (!loadedTuples.has(tuple)) continue;
    metrics[weight === 400 && style === 'normal' ? family : tuple] = entry.metric;
  }
  return { faces: loadedFaces, metrics: Object.freeze(metrics) };
}

/** bold / boldItalic slots ⇒ CSS `font-weight: bold`; otherwise `normal`. */
function weightForStyle(style: EmbeddedFontRef['style']): 'normal' | 'bold' {
  return style === 'bold' || style === 'boldItalic' ? 'bold' : 'normal';
}

/** italic / boldItalic slots ⇒ CSS `font-style: italic`; otherwise `normal`. */
function styleForStyle(style: EmbeddedFontRef['style']): 'normal' | 'italic' {
  return style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal';
}
