import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { CjkLang, OfficeFontFallbackRequest } from '@silurus/ooxml-core';
import { findReferenceFontMetrics } from '@silurus/ooxml-core';

// Kept in a parser-model-free module: layout font routing reads the declared
// scopes, while this module traverses parser-owned facts.
export { declaredFontSubstitutions, scriptScopedSubstitutedFamilies } from './font-substitution-scope.js';
import type {
  DocxDocumentModel,
} from './types.js';
import { docxRenderedTextUsages } from './document-content.js';

/**
 * DOCX font requests every load makes: the exact local Office-face tuples and
 * family presence probes rendered text needs, and the script CJK language.
 * The font-source demand plan (`LoadOptions.fonts` only) is
 * font-demand-plan.ts.
 */

function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  for (const usage of docxRenderedTextUsages(doc)) yield usage.text;
}

/** Probe exact local style tuples used by rendered text. The shared loader
 * declines uncatalogued names, and the document font table alone never queues
 * a face that rendered content does not use. No font bytes are packaged.
 * `fontDemand` (font sources configured) collects through the same
 * `'font-demand'` projection as `docxFontDemandCandidates` (font-demand-plan.ts), so every
 * demanded family also receives its presence probe. */
export function docxOfficeFontFallbackRequests(
  doc: DocxDocumentModel,
  options: Readonly<{ fontDemand?: boolean; cjkFallback?: CjkLang }> = {},
): OfficeFontFallbackRequest[] {
  const tuples = new Map<string, OfficeFontFallbackRequest>();
  const add = (family: string | null | undefined, bold = false, italic = false) => {
    const name = family?.trim();
    if (!name) return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    tuples.set(`${name.toLowerCase()}:${weight}:${style}`, { family: name, weight, style });
  };
  // Theme faces can govern paragraph marks and inherited runs even when a
  // particular authored run does not repeat its font name.
  add(doc.majorFont);
  add(doc.minorFont);
  for (const usage of docxRenderedTextUsages(doc, options.fontDemand ? 'font-demand' : 'native-preflight')) {
    for (const family of usage.fontFamilies) add(family, usage.bold, usage.italic);
    if (usage.text && (usage.latinFontFamily === null ||
      (usage.latinFontFamily === undefined && !usage.fontFamilies.some(Boolean)))) {
      // A Latin slot can inherit the theme even when another script slot has an
      // authored face. The inherited bold/italic axes need their own resource.
      add(doc.minorFont, usage.bold, usage.italic);
    }
  }
  // LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY compares a styled run against the
  // regular face it may be synthesized from, so a styled tuple the catalog
  // lacks also loads its family's catalogued regular face.
  for (const request of [...tuples.values()]) {
    const weight = request.weight ?? 400;
    const style = request.style ?? 'normal';
    if (weight === 400 && style === 'normal') continue;
    if (findReferenceFontMetrics(request.family, { weight, style }).length > 0) continue;
    if (findReferenceFontMetrics(request.family, { weight: 400, style: 'normal' }).length === 0) continue;
    add(request.family);
  }
  // Every authored family is checked for installation as a family, independent
  // of the styles the document uses: an installed family paints every tuple
  // (layout's installedFamilies) and is never demanded from a font source
  // (font-demand-plan.ts docxFontFamilyPlan).
  const presence = new Map<string, OfficeFontFallbackRequest>();
  const addPresence = (family: string) => {
    const key = family.trim().toLowerCase();
    if (key && !presence.has(key)) presence.set(key, { family: family.trim(), presenceOnly: true });
  };
  for (const request of tuples.values()) addPresence(request.family);
  return [...tuples.values(), ...presence.values()];
}


export function docxScriptCjkLanguage(doc: DocxDocumentModel): CjkLang | null {
  const scripts = new ScriptPreloadAccumulator(null);
  scripts.addText(docxTextRuns(doc));
  return scripts.scriptCjkLanguage();
}
