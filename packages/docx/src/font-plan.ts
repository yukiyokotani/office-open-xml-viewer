import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { CjkLang, FontSubstituteScript, OfficeFontFallbackRequest } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkLangFromLanguage,
  findReferenceFontMetrics,
  scriptPreloadNamesForText,
  substituteEntryCoversText,
} from '@silurus/ooxml-core';
import {
  chartFontFamilies,
  fontFamilyKey,
  probeFontFamilyPresence,
  substitutionsFor,
  type FontFamilyPresence,
  type FontSource,
  type FontSubstitution,
} from '@silurus/ooxml-core/internal/font-sources';
import { declaredFontSubstitutions } from './font-substitution-scope.js';

// Kept in a parser-model-free module: layout font routing reads the declared
// scopes, while this module traverses parser-owned facts.
export { declaredFontSubstitutions, scriptScopedSubstitutedFamilies } from './font-substitution-scope.js';
import type {
  DocxDocumentModel,
} from './types.js';
import { docxRenderedTextUsages } from './document-content.js';

/**
 * DOCX font demand: which families a document asks the configured font
 * sources (`LoadOptions.fonts`) for. Library resource policy: ECMA-376 §17.8
 * names faces and leaves their provision to the host.
 *
 * Demand is computed from the parsed model alone. Which families a source
 * actually serves is the source's business (an allow-list, a stylesheet
 * index, a resolver), so no provider table appears here. Main-thread and
 * worker loads compute it from the same model with the same function, so they
 * demand, register and load identical faces (worker/main pixel equivalence).
 * The demanded Unicode scalars are collected separately, in the lazily loaded
 * `font-source-load.ts`, so the default (no source) path never loads the
 * collector's tables.
 */

/** Every chart model of the document (§17.3.3.9 drawing → c:chart), from
 * one generic walk over the parsed model. */
function* docxChartModels(doc: DocxDocumentModel): Generator<unknown> {
  const pending: unknown[] = [doc];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value) || ArrayBuffer.isView(value)) continue;
    seen.add(value);
    const record = value as Record<string, unknown>;
    if (record.type === 'chart' && record.chart && typeof record.chart === 'object') {
      yield record.chart;
      continue;
    }
    for (const child of Object.values(record)) if (child !== null && typeof child === 'object') pending.push(child);
  }
}

function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  for (const usage of docxRenderedTextUsages(doc)) yield usage.text;
}

/** Families a document may paint, before the installed-face check. */
export interface DocxFontDemandCandidates {
  /** Theme major/minor, then every rendered authored family, first spelling
   * per family. */
  readonly authored: readonly string[];
  /** Library script fallback families (Noto per script present in the text,
   * CJK ordered by language). */
  readonly fallback: readonly string[];
  /** Per authored family (lower-cased), the scripts of its declared
   * script-scoped substitutes that its rendered text proves
   * (`FontSessionDemand.substituteScripts`). */
  readonly substituteScripts: Readonly<Record<string, readonly FontSubstituteScript[]>>;
}

/**
 * The families to demand for a document: the theme major/minor fonts, every
 * family rendered text resolves to, plus only the script-fallback faces whose
 * script the document's TEXT actually contains ({@link scriptPreloadNamesForText}).
 *
 * Rendered families come from {@link docxRenderedTextUsages} in its
 * `'font-demand'` projection: all four §17.3.2.26 slots (ascii, hAnsi,
 * eastAsia, cs) resolved through the shaper's own slot rule, across body,
 * tables, headers/footers, notes, complete text-box stories and numbering
 * markers. The document font table alone never adds a name.
 *
 * Every authored family is demanded, so a source serving it under its own
 * name is always consulted (a face of the authored family beats a
 * substitute, design §3.3). A script-scoped substitute (Sakkal Majalla →
 * Noto Naskh Arabic, Arabic only) may paint only text of its script, so
 * `substituteScripts` records, per family, the scoped scripts its rendered
 * text proves; the session consults a scoped substitution only for those, and
 * never requests a substitute nothing would paint. Theme fonts govern
 * paragraph marks and inherited runs whatever their text, so every declared
 * scoped script is admitted for them.
 *
 * Eagerly fetching multi-MB CJK families for a document with no CJK glyphs
 * would block first paint for nothing, so script fallbacks follow the text.
 */
export function docxFontDemandCandidates(
  doc: DocxDocumentModel,
  substitutes: readonly FontSubstitution[],
  fallback?: CjkLang,
): DocxFontDemandCandidates {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? fallback ?? null;
  const scripts = new ScriptPreloadAccumulator(cjkLang);
  const languageNames = new Set<string>();
  const authored = new Map<string, string>();
  const addAuthored = (family: string | null | undefined) => {
    const name = family?.trim();
    const key = name ? fontFamilyKey(name) : '';
    if (name && key && !authored.has(key)) authored.set(key, name);
  };
  const scopedScripts = new Map<string, Set<FontSubstituteScript>>();
  for (const substitution of substitutes) {
    if (!substitution.script) continue;
    const key = fontFamilyKey(substitution.for);
    scopedScripts.set(key, (scopedScripts.get(key) ?? new Set()).add(substitution.script));
  }
  const proven = new Map<string, Set<FontSubstituteScript>>();
  const prove = (key: string, script: FontSubstituteScript) =>
    proven.set(key, (proven.get(key) ?? new Set()).add(script));
  for (const theme of [doc.majorFont, doc.minorFont]) {
    const name = theme?.trim();
    if (!name) continue;
    addAuthored(name);
    for (const script of scopedScripts.get(fontFamilyKey(name)) ?? []) prove(fontFamilyKey(name), script);
  }
  for (const usage of docxRenderedTextUsages(doc, 'font-demand')) {
    for (const family of usage.fontFamilies) {
      const name = family?.trim();
      if (!name) continue;
      addAuthored(name);
      const key = fontFamilyKey(name);
      for (const script of scopedScripts.get(key) ?? []) {
        if (!proven.get(key)?.has(script) && substituteEntryCoversText({ script }, usage.text, 'any')) prove(key, script);
      }
    }
    const region = cjkLangFromLanguage(usage.eastAsiaLanguage);
    if (region) {
      for (const name of scriptPreloadNamesForText([usage.text], region, true)) languageNames.add(name);
    } else {
      scripts.addText([usage.text]);
    }
  }
  // Chart text paints its own faces (title, axes, labels, legend, theme
  // faces) with the loaded unscoped substitutions (ChartModel
  // .fontSubstitutions), so a face named only by a chart is demanded too.
  for (const chart of docxChartModels(doc)) for (const family of chartFontFamilies(chart)) addAuthored(family);
  return {
    authored: [...authored.values()],
    fallback: [...new Set([...scripts.names(), ...languageNames])],
    substituteScripts: Object.freeze(Object.fromEntries([...authored.keys()].map((key) =>
      [key, Object.freeze([...(proven.get(key) ?? [])].sort())]))),
  };
}

/** The family part of one document's font demand. */
export interface DocxFontFamilyPlan {
  /** Authored families not proven installed, in demand order. */
  readonly families: readonly string[];
  /** Library script fallback families. */
  readonly fallbackFamilies: readonly string[];
  /** See {@link DocxFontDemandCandidates.substituteScripts}. */
  readonly substituteScripts: Readonly<Record<string, readonly FontSubstituteScript[]>>;
  /** Normalised candidate families with a declared substitute whose own face
   * is installed. They are neither demanded nor routed to the substitute, and
   * their script scope is lifted (an installed authored face paints all of
   * its text). */
  readonly installedSubstituteFamilies: readonly string[];
}

/**
 * Library policy: the authored family names the face to use, so an installed
 * authored face is never displaced by a source face. A family counts as
 * installed when any exact local route of it loaded, or when the family
 * presence probe that {@link docxOfficeFontFallbackRequests} queues for every
 * authored family found it. Installed families are therefore never demanded,
 * and a source is never asked for a family the host already has.
 */
export function docxFontFamilyPlan(
  doc: DocxDocumentModel,
  sources: readonly FontSource[],
  cjkFallback: CjkLang | undefined,
  officeFonts: DocxOfficeFontPresence,
): DocxFontFamilyPlan {
  const substitutes = declaredFontSubstitutions(sources);
  const candidates = docxFontDemandCandidates(doc, substitutes, cjkFallback);
  const probed = new Set(officeFonts.installed ?? []);
  const checked = officeFonts.presenceChecked === undefined ? undefined : new Set(officeFonts.presenceChecked);
  // Any proven face of the family counts, whichever styles produced it.
  const routedFamilies = new Set(Object.keys(officeFonts.routes).map((key) => key.split(':')[0]));
  const installed = (family: string) => {
    const key = family.trim().toLowerCase();
    return routedFamilies.has(key) || probed.has(key);
  };
  // Only a family proven missing is demanded: one whose presence probe did
  // not complete (past the probe's source cap or deadline) keeps its host
  // face, so a source face never shadows an installed font it was not
  // checked against. Without presence facts every family counts as checked.
  const families = candidates.authored.filter((family) => !installed(family)
    && (checked === undefined || checked.has(family.trim().toLowerCase())));
  // A family with an admissible substitute: unscoped, or scoped to a script
  // its text proves. Only such a substitute could have displaced it.
  // Own-property read: keys are document font names (e.g. `constructor`).
  const scriptsOf = (family: string): readonly FontSubstituteScript[] => {
    const key = fontFamilyKey(family);
    return Object.prototype.hasOwnProperty.call(candidates.substituteScripts, key)
      ? candidates.substituteScripts[key]
      : [];
  };
  const substituted = (family: string) => substitutionsFor(substitutes, family, scriptsOf(family)).length > 0;
  const installedSubstituteFamilies = [...new Set(candidates.authored
    .filter((family) => installed(family) && substituted(family))
    .map((family) => family.trim().toLowerCase()))].sort();
  return Object.freeze({
    families: Object.freeze(families),
    fallbackFamilies: Object.freeze([...candidates.fallback]),
    installedSubstituteFamilies: Object.freeze(installedSubstituteFamilies),
    substituteScripts: candidates.substituteScripts,
  });
}

/** The installed-face facts {@link docxFontFamilyPlan} reads: exact local
 * routes, families the presence probe found installed, and every family whose
 * presence probe completed. */
export type DocxOfficeFontPresence = Readonly<{
  routes: Readonly<Record<string, unknown>>;
  installed?: readonly string[];
  presenceChecked?: readonly string[];
}>;

/**
 * Complete the presence facts of `officeFonts` for font-source demand: the
 * office-fallback load probes at most one capped batch of families, so every
 * authored candidate it did not prove installed or check is probed now in
 * batches ({@link probeFontFamilyPresence}), and the results are merged. Only
 * called with font sources configured; the default path keeps its single
 * capped probe.
 */
export async function completeDocxFontPresence<T extends DocxOfficeFontPresence>(
  doc: DocxDocumentModel,
  sources: readonly FontSource[],
  cjkFallback: CjkLang | undefined,
  officeFonts: T,
  probe: (families: readonly string[]) => Promise<FontFamilyPresence> = (families) => probeFontFamilyPresence(families),
): Promise<T> {
  const routed = new Set(Object.keys(officeFonts.routes).map((key) => key.split(':')[0]));
  const installed = new Set(officeFonts.installed ?? []);
  const checked = new Set(officeFonts.presenceChecked ?? []);
  const unproven = docxFontDemandCandidates(doc, declaredFontSubstitutions(sources), cjkFallback).authored
    .filter((family) => {
      const key = family.trim().toLowerCase();
      return !routed.has(key) && !installed.has(key) && !checked.has(key);
    });
  if (unproven.length > 0) {
    const presence = await probe(unproven);
    for (const family of presence.installed) installed.add(family);
    for (const family of presence.checked) checked.add(family);
  }
  return {
    ...officeFonts,
    ...(installed.size > 0 ? { installed: [...installed].sort() } : {}),
    presenceChecked: [...checked].sort(),
  };
}

/** Probe exact local style tuples used by rendered text. The shared loader
 * declines uncatalogued names, and the document font table alone never queues
 * a face that rendered content does not use. No font bytes are packaged.
 * `fontDemand` (font sources configured) collects through the same
 * `'font-demand'` projection as {@link docxFontDemandCandidates}, so every
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
  // (docxFontFamilyPlan).
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
