import {
  fontFamilyKey,
  googleFonts,
  substitutionsFor,
  type FontSource,
  type FontSubstitution,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import type { OpenTypeLineMetrics } from '@silurus/ooxml-core';
import { declaredFontSubstitutions, docxFontFamilyPlan } from '../font-plan.js';
import type { DocxDocumentModel } from '../types.js';

type LoadedFaceRecord = Readonly<{
  family: string; weight?: string; style?: string; status?: string;
  unicodeRange?: string;
  /** Parsed tables of the face's bytes, as a session snapshots them. */
  resourceMetrics?: OpenTypeLineMetrics;
}>;

/**
 * Layout-service options for a unit test that stands in for a completed
 * font-source load: the document's real demand plan (`docxFontFamilyPlan`)
 * against `sources`, with `faces` as the faces that loaded. Each demanded
 * family is claimed the way a session claims it: by a loaded face of its own
 * family, else by the first declared substitute whose face loaded.
 */
export function sourceFontLayoutOptions(
  model: DocxDocumentModel,
  faces: readonly LoadedFaceRecord[],
  options: Readonly<{ installed?: readonly string[]; sources?: readonly FontSource[] }> = {},
): {
  fontSubstitutes: readonly FontSubstitution[];
  sourceFonts: { names: readonly string[]; loaded: LoadedFontSources };
  installedSubstituteFamilies: readonly string[];
} {
  const sources = options.sources ?? [googleFonts()];
  const plan = docxFontFamilyPlan(model, sources, undefined, { routes: {}, installed: [...(options.installed ?? [])] });
  const substitutes = declaredFontSubstitutions(sources);
  const names = [...plan.families, ...plan.fallbackFamilies];
  const loadedFaces = faces.filter((face) => (face.status ?? 'loaded') === 'loaded');
  const loadedFamilies = new Set(loadedFaces.map((face) => fontFamilyKey(face.family)));
  const substitutions = new Map<string, { family: string; script?: FontSubstitution['script'] }>();
  for (const name of names) {
    const key = fontFamilyKey(name);
    if (loadedFamilies.has(key)) continue;
    const substitution = substitutionsFor(substitutes, name, plan.substituteScripts[key] ?? []).find((entry) => loadedFamilies.has(fontFamilyKey(entry.use)));
    if (substitution) {
      substitutions.set(key, { family: substitution.use, ...(substitution.script ? { script: substitution.script } : {}) });
    }
  }
  return {
    fontSubstitutes: substitutes,
    installedSubstituteFamilies: plan.installedSubstituteFamilies,
    sourceFonts: {
      names,
      loaded: {
        faces: loadedFaces.map((face) => ({
          family: face.family,
          weight: face.weight ?? '400',
          style: face.style ?? 'normal',
          stretch: 'normal',
          unicodeRange: face.unicodeRange ?? 'U+0-10FFFF',
          sourceIndex: 0,
          face: face as FontFace,
          ...(face.resourceMetrics ? { resourceMetrics: face.resourceMetrics } : {}),
        })),
        substitutions,
        regularAliases: new Map(),
        // As a session registers them: one private family per served family.
        families: new Map([...loadedFamilies].map((key, index) => [key, `ooxml-font-source-test-${index}`])),
      },
    },
  };
}
