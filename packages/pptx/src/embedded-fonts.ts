import {
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type EmbeddedFontFace,
} from '@silurus/ooxml-core';
import type { PptxEmbeddedFontRef } from './worker-protocol';

export interface LoadedPptxEmbeddedFonts {
  readonly faces: FontFace[];
  /** Lower-cased authored family → presentation-scoped FontFace family. */
  readonly aliases: ReadonlyMap<string, string>;
  /** Presentation-scoped FontFace family → lower-cased authored family. */
  readonly authoredFamilies: ReadonlyMap<string, string>;
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
  const batchSize = 2;
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
        } catch {
          return null;
        }
      },
    ));
    const loadable = faces.filter((face): face is EmbeddedFontFace => face !== null);
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
  const loadedAliases = new Set(loaded.map((face) => normalizedFamily(face.family)));
  const aliases = new Map(
    [...candidateAliases].filter(([, alias]) => loadedAliases.has(normalizedFamily(alias))),
  );
  const authoredFamilies = new Map(
    [...aliases].map(([authored, alias]) => [alias, authored]),
  );
  return { faces: loaded, aliases, authoredFamilies };
}

/** Do not register a web substitute for a family successfully loaded from the deck. */
export function excludeEmbeddedFontFamilies(
  names: readonly (string | null)[],
  loadedAliases: ReadonlyMap<string, string>,
): (string | null)[] {
  return names.filter((name) => name === null || !loadedAliases.has(name.trim().toLowerCase()));
}
