import {
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type EmbeddedFontFace,
  type OfficeFontFallbackRequest,
  isSymbolFontFamily,
  symbolCmapEncoding,
  symbolFaceKey,
  type SymbolFaceEncoding,
} from '@silurus/ooxml-core';
import type { LoadedFontSources } from '@silurus/ooxml-core/internal/font-sources';
import { parseResourceLimitError } from '@silurus/ooxml-core/worker';
import type { PptxEmbeddedFontRef } from './worker-protocol';

export interface LoadedPptxEmbeddedFonts {
  readonly faces: FontFace[];
  /** Lower-cased authored family → presentation-scoped FontFace family. */
  readonly aliases: ReadonlyMap<string, string>;
  /** Presentation-scoped FontFace family → lower-cased authored family. */
  readonly authoredFamilies: ReadonlyMap<string, string>;
  /** Successfully registered authored family/style slots (§19.2.1.9). */
  readonly tuples: ReadonlySet<string>;
  /** {@link symbolFaceKey} of a presentation-scoped face → the symbol codes
   *  its bytes draw, for symbol-font families (core symbol-face.ts). */
  readonly symbolEncodings: ReadonlyMap<string, SymbolFaceEncoding>;
}

/**
 * Symbol-font faces whose bytes this presentation holds, by
 * {@link symbolFaceKey}: embedded faces (one per style slot, §19.2.1.9), and
 * served families whose faces CSS matching can pick for a tuple all have read
 * cmaps (LoadedFontSources.symbolEncodings). The renderer prefers these cmap
 * readings over measuring (renderer symbolFaceSegments).
 */
export function pptxSymbolFaceEncodings(
  embedded: ReadonlyMap<string, SymbolFaceEncoding> | undefined,
  served: LoadedFontSources | undefined,
): ReadonlyMap<string, SymbolFaceEncoding> | undefined {
  const out = new Map([...(embedded ?? []), ...(served?.symbolEncodings ?? [])]);
  return out.size > 0 ? out : undefined;
}

/** Only an actually registered PresentationML face occupies its style slot. */
export function uncoveredOfficeFontRequests(
  requests: readonly OfficeFontFallbackRequest[],
  embeddedTuples: ReadonlySet<string>,
): OfficeFontFallbackRequest[] {
  return requests.filter((request) =>
    !embeddedTuples.has(`${request.family.toLowerCase()}:${request.weight}:${request.style}`));
}

let nextFontScope = 1;

function normalizedFamily(value: string): string {
  return value.trim().toLowerCase();
}

/**
 * Load PresentationML font parts and register them before text measurement.
 * PPTX font parts are raw sfnt or EOT (ECMA-376 Part 1 §15.2.13), never the
 * WordprocessingML ODTTF obfuscation format.
 */
export async function loadEmbeddedFonts(
  refs: readonly PptxEmbeddedFontRef[],
  fetchFontBytes: (partPath: string) => Promise<Uint8Array>,
): Promise<LoadedPptxEmbeddedFonts> {
  if (refs.length === 0) return {
    faces: [],
    aliases: new Map(),
    authoredFamilies: new Map(),
    tuples: new Set(),
    symbolEncodings: new Map(),
  };
  const scope = nextFontScope++;
  const candidateAliases = new Map<string, string>();
  for (const ref of refs) {
    const key = normalizedFamily(ref.fontName);
    if (!candidateAliases.has(key)) {
      candidateAliases.set(key, `__ooxml_pptx_${scope}_${candidateAliases.size + 1}`);
    }
  }
  // Retain no more than two unregistered WASM/transfer buffers at once. Each
  // batch is copied into FontFace storage before the next extraction begins.
  const loaded: FontFace[] = [];
  const held = new Set<FontFace>();
  const symbolCandidates = new Map<string, SymbolFaceEncoding>();
  const batchSize = 2;
  try {
    for (let offset = 0; offset < refs.length; offset += batchSize) {
      const faces = await Promise.all(refs.slice(offset, offset + batchSize).map(
        async (ref): Promise<EmbeddedFontFace | null> => {
          try {
            return {
              family: candidateAliases.get(normalizedFamily(ref.fontName)) as string,
              bytes: await fetchFontBytes(ref.partPath),
              odttf: false,
              weight: ref.style === 'bold' || ref.style === 'boldItalic' ? 'bold' : 'normal',
              style: ref.style === 'italic' || ref.style === 'boldItalic' ? 'italic' : 'normal',
            };
          } catch (error) {
            // Missing/malformed fonts are optional; archive/resource policy is not.
            if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
            // Worker rendering reads WASM directly, before bridge deserialization.
            const resourceError = parseResourceLimitError(error);
            if (resourceError) throw resourceError;
            return null;
          }
        },
      ));
      const loadable = faces.filter((face): face is EmbeddedFontFace => face !== null);
      for (const face of loadable) {
        const authored = refs.find((ref) => candidateAliases.get(normalizedFamily(ref.fontName)) === face.family)?.fontName;
        const key = symbolFaceKey(face.family, face.weight === 'bold' ? 700 : 400, face.style === 'italic' ? 'italic' : 'normal');
        if (!isSymbolFontFamily(authored) || symbolCandidates.has(key)) continue;
        // Raw sfnt only (an EOT part has no readable cmap here; it is measured).
        const encoding = symbolCmapEncoding(face.bytes);
        if (encoding) symbolCandidates.set(key, encoding);
      }
      if (loadable.length === 0) continue;
      for (const face of await registerEmbeddedFonts(loadable)) {
        if (held.has(face)) {
          // A content-identical face retained by an earlier batch needs no second
          // holder from this presentation. Balance that registry retain now.
          unregisterEmbeddedFonts([face]);
        } else {
          held.add(face);
          loaded.push(face);
        }
      }
    }
  } catch (error) {
    // The caller cannot own a rejected result. Balance only this load's retains,
    // including faces registered before a later extraction batch crossed a cap.
    unregisterEmbeddedFonts(loaded);
    throw error;
  }
  const loadedAliases = new Set(loaded.map((face) => normalizedFamily(face.family)));
  const aliases = new Map(
    [...candidateAliases].filter(([, alias]) => loadedAliases.has(normalizedFamily(alias))),
  );
  const authoredFamilies = new Map(
    [...aliases].map(([authored, alias]) => [alias, authored]),
  );
  // An embedded regular face does not satisfy the bold or italic slots. Keep
  // this distinct from the family alias so a missing Calibri slot may use its
  // measured Office fallback without replacing the successful embedded face.
  const tuples = new Set(loaded.map((face) => {
    const authored = authoredFamilies.get(face.family) as string;
    return `${authored}:${face.weight === 'bold' || face.weight === '700' ? 700 : 400}:${face.style === 'italic' ? 'italic' : 'normal'}`;
  }));
  // Only a style slot whose face registered paints its bytes; a failed slot
  // is painted by another face (CSS matching), which the renderer measures.
  const registered = new Set(loaded.map((face) => symbolFaceKey(face.family,
    face.weight === 'bold' || face.weight === '700' ? 700 : 400, face.style === 'italic' ? 'italic' : 'normal')));
  const symbolEncodings = new Map([...symbolCandidates].filter(([key]) => registered.has(key)));
  return { faces: loaded, aliases, authoredFamilies, tuples, symbolEncodings };
}

/** Do not register a web substitute for a family successfully loaded from the deck. */
export function excludeEmbeddedFontFamilies(
  names: readonly (string | null)[],
  loadedAliases: ReadonlyMap<string, string>,
): (string | null)[] {
  return names.filter((name) => name === null || !loadedAliases.has(name.trim().toLowerCase()));
}
