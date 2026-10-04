import { parseOpenTypeResourceWithSupport } from '@silurus/ooxml-core/internal/open-type-resource-support';
import { fontSupportFacts, type FontSupportFacts } from '@silurus/ooxml-core/internal/font-support-registry';
import {
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type EmbeddedFontFace,
  type OfficeFontFallbackRequest,
  type OpenTypeLineMetrics,
} from '@silurus/ooxml-core';
import type { PptxEmbeddedFontRef } from './worker-protocol';
import { powerPointResourceFaceMetrics, type PowerPointFaceMetrics } from './powerpoint-line-metrics.js';

/** A style slot may contain several subset resources (§19.2.1.9/.10).
 * Entries follow FontFaceSet insertion order, including unreadable resources;
 * undefined is an ownership barrier, not permission to borrow another part. */
export interface PptxEmbeddedFontResource {
  readonly identity: string;
  readonly metrics: PowerPointFaceMetrics | undefined;
  readonly tables: OpenTypeLineMetrics | undefined;
  readonly support: FontSupportFacts | undefined;
}
export type PptxEmbeddedFontMetrics = ReadonlyMap<string, readonly PptxEmbeddedFontResource[]>;

export interface LoadedPptxEmbeddedFonts {
  readonly faces: FontFace[];
  /** Lower-cased authored family → presentation-scoped FontFace family. */
  readonly aliases: ReadonlyMap<string, string>;
  /** Presentation-scoped FontFace family → lower-cased authored family. */
  readonly authoredFamilies: ReadonlyMap<string, string>;
  /** Successfully registered authored family/style slots (§19.2.1.9). */
  readonly tuples: ReadonlySet<string>;
  /** Resource metrics/cmaps per authored tuple, in successful insertion order.
   * An unsupported part (e.g. EOT) retains an undefined entry when it loads. */
  readonly metrics: PptxEmbeddedFontMetrics;
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
    metrics: new Map(),
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
  // FontFace identity, rather than family/style, binds each successful
  // registration to its own tables. CT_EmbeddedFontList has no uniqueness
  // constraint; distinct same-slot parts can cover different glyph subsets.
  const parsed = new Map<FontFace, PptxEmbeddedFontResource>();
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
    const registrations = await Promise.all(loadable.map(async (resource) => {
      const tables = parseOpenTypeResourceWithSupport(resource.bytes);
      const metrics = tables ? powerPointResourceFaceMetrics(tables) : undefined;
      // Register separately to retain the input-resource → FontFace binding
      // even when a sibling fails. Calls add synchronously in input order;
      // only their load/readiness waits run concurrently (at most two).
      return { faces: await registerEmbeddedFonts([resource]), metrics, tables: tables ?? undefined, support: fontSupportFacts(tables ?? undefined) };
    }));
    for (const registration of registrations) {
      for (const face of registration.faces) {
        if (held.has(face)) {
          // A reused face was not reinserted into FontFaceSet. Preserve its
          // original position/metrics and balance this batch's extra retain.
          unregisterEmbeddedFonts([face]);
        } else {
          held.add(face);
          loaded.push(face);
          parsed.set(face, Object.freeze({ identity: `pptx:${scope}:${loaded.length}`,
            metrics: registration.metrics, tables: registration.tables, support: registration.support }));
        }
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
  // An embedded regular face does not satisfy the bold or italic slots. Keep
  // this distinct from the family alias so a missing Calibri slot may use its
  // measured Office fallback without replacing the successful embedded face.
  const tuples = new Set(loaded.map((face) => {
    const authored = authoredFamilies.get(face.family) as string;
    return `${authored}:${face.weight === 'bold' || face.weight === '700' ? 700 : 400}:${face.style === 'italic' ? 'italic' : 'normal'}`;
  }));
  const metrics = new Map<string, PptxEmbeddedFontResource[]>();
  for (const face of loaded) {
    const authored = authoredFamilies.get(face.family) as string;
    const key = `${authored}:${face.weight === 'bold' || face.weight === '700' ? 700 : 400}:${face.style === 'italic' ? 'italic' : 'normal'}`;
    const resources = metrics.get(key) ?? [];
    resources.push(parsed.get(face) as PptxEmbeddedFontResource);
    metrics.set(key, resources);
  }
  return { faces: loaded, aliases, authoredFamilies, tuples, metrics };
}

/** Do not register a web substitute for a family successfully loaded from the deck. */
export function excludeEmbeddedFontFamilies(
  names: readonly (string | null)[],
  loadedAliases: ReadonlyMap<string, string>,
): (string | null)[] {
  return names.filter((name) => name === null || !loadedAliases.has(name.trim().toLowerCase()));
}
