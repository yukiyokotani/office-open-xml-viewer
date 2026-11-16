/**
 * DOCX font-source loading (`LoadOptions.fonts`): demanded Unicode scalars and
 * the per-document {@link FontSourceSession}. Library resource policy, not
 * OOXML behaviour: ECMA-376 §17.8 names faces and leaves provision to the host.
 *
 * Imported lazily (main thread) only when sources are configured, so the
 * default offline path never loads the session, the CSS adapters or the
 * inert-scalar tables. The main thread and the render worker run this same
 * module against the same parsed model with the same descriptors, so both
 * realms demand, register and load identical faces.
 */
import {
  FontDemandScalarCollector,
  VERTICAL_PRESENTATION_FORMS,
  type FontDemandScalars,
} from '@silurus/ooxml-core/internal/font-scalar-demand';
import {
  FontSourceSession,
  readFontResourceMetrics,
  type FontSourceSessionOptions,
  type FontSourceDescriptor,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import {
  isMetafileMime,
  isSymbolFontFamily,
  parseDateTimePictureSwitch,
  parseFieldFormatSwitch,
} from '@silurus/ooxml-core';
import { docxRenderedTextUsages } from './document-content.js';
import type { DocxFontFamilyPlan } from './font-demand-plan.js';
import { servedFontLineMetrics } from './served-font-metrics.js';
import type { DocxSourceFonts } from './source-font-metrics.js';

// The demand plan travels with the session: the main thread imports both
// together, and only when sources are configured.
export { completeDocxFontPresence, docxFontFamilyPlan } from './font-demand-plan.js';
import { emphasisGlyph } from './layout/paragraph.js';
import { LAYOUT_SYNTHESIZED_SCALARS } from './line-layout.js';
import { mathFallbackText } from './layout/math-fallback-text.js';
import type { DocxDocumentModel } from './types.js';

/** ST_NumberFormat values (§17.18.59) whose `formatOrdinalNumber` output is
 * ASCII for every value: decimal and its ASCII variants, roman, Latin letters,
 * hex, the dashed decimal, and the spell-out formats the converter renders as
 * decimal (documented residual in core number-format.ts). Any other format can
 * emit scalars no text spells (CJK numerals, Hebrew, Thai digits, …), so it
 * makes demand `'all'`. */
export const ASCII_NUMBER_FORMATS: ReadonlySet<string> = new Set([
  'decimal', 'decimalHalfWidth', 'decimalZero', 'upperRoman', 'lowerRoman',
  'upperLetter', 'lowerLetter', 'hex', 'numberInDash', 'none',
  'cardinalText', 'ordinalText', 'ordinal',
]);

/** §17.18.93 ST_TextDirection, strict and transitional spellings: `lrTb`
 * is horizontal; every other value lays the text out vertically
 * (vertical-text.ts), which substitutes vertical presentation forms. */
const TEXT_DIRECTIONS: ReadonlySet<string> = new Set([
  'lrTb', 'tbRl', 'btLr', 'lrTbV', 'tbRlV', 'tbLrV', 'tb', 'rl', 'lr', 'tbV', 'rlV', 'lrV',
]);

const METAFILE_PATH = /\.(?:wmf|emf|wmz|emz)$/i;

/** The CSS weights a WordprocessingML run paints with: regular and bold. */
const DOCX_RUN_WEIGHTS: readonly number[] = Object.freeze([400, 700]);

/**
 * Paint facts the rendered-text traversal does not spell, from one walk over
 * the parsed model (the same generic walk `vertical-render-capability.ts`
 * uses). Returns the extra texts layout synthesises for this document, or
 * `'all'` for a paint path whose scalars cannot be proven from the text:
 *
 * - text-box direction other than `horz` and §17.3.2.10 `eastAsianLayout`
 *   vert/combine (combined brackets are generated text);
 * - an unknown section or cell text direction;
 * - charts (their own text and number formatting);
 * - page / note numbers in a non-ASCII number format.
 *
 * WMF/EMF pictures need no load-time demand: their text records demand
 * their own scalars when they play (core `bindMetafileFontDemand`).
 *
 * Synthesised texts: `allCaps` / `smallCaps` runs paint their text upper-cased
 * (§17.3.2.5/.33), emphasis marks paint their glyph (§17.3.2.12), a DATE /
 * TIME picture (§17.16.4.1) paints its literal characters, a ruby run
 * (§17.3.3.25) paints its guide text, an OMML equation (§22.1) that no math
 * renderer converted is sized by measuring its deterministic text fallback
 * (layout math-fallback-text.ts: the run's own characters, operators and the
 * fallback's `/^_√(),` punctuation; a converted equation paints a raster
 * whose glyphs reach no document font), and a vertical section or cell
 * (§17.6.20 / §17.4.72 textDirection) paints the vertical presentation forms
 * vertical-text.ts substitutes, and kashida justification paints U+0640.
 * Right-to-left runs and bidi paragraphs paint their texts reordered; the
 * core collector demands what the shaper may substitute (mirrored partners,
 * Arabic presentation forms).
 */
function docxSynthesizedPaint(doc: DocxDocumentModel): 'all' | string[] {
  const texts: string[] = [];
  let verticalForms = false;
  let kashida = false;
  const settings = (doc as { __noteLayoutSettings?: { footnoteNumberFormat?: string; endnoteNumberFormat?: string } })
    .__noteLayoutSettings;
  for (const format of [settings?.footnoteNumberFormat, settings?.endnoteNumberFormat]) {
    if (format !== undefined && !ASCII_NUMBER_FORMATS.has(format)) return 'all';
  }
  const pending: unknown[] = [doc];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value)) continue;
    seen.add(value);
    if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) continue;
    if (!Array.isArray(value)) {
      const record = value as Record<string, unknown>;
      if (typeof record.textDirection === 'string') {
        if (!TEXT_DIRECTIONS.has(record.textDirection)) return 'all';
        if (record.textDirection !== 'lrTb' && !verticalForms) {
          verticalForms = true;
          texts.push(VERTICAL_PRESENTATION_FORMS);
        }
      }
      if (typeof record.textVert === 'string' && record.textVert !== 'horz') return 'all';
      if (record.eastAsianVert === true || record.eastAsianCombine === true) return 'all';
      // §17.18.44 kashida justification elongates Arabic words with inserted
      // U+0640 TATWEEL (kashida-justify.ts).
      if (typeof record.alignment === 'string' && record.alignment.endsWith('Kashida') && !kashida) {
        kashida = true;
        texts.push('\u0640');
      }
      if (record.type === 'chart') return 'all';
      if (record.type === 'math') {
        texts.push(typeof record.fallbackText === 'string'
          ? record.fallbackText : mathFallbackText(record.nodes as Parameters<typeof mathFallbackText>[0]));
        continue;
      }
      // §17.6.12 pgNumType fmt (body and per-section markers).
      if (record.pageNumType && typeof record.pageNumType === 'object') {
        const fmt = (record.pageNumType as { fmt?: unknown }).fmt;
        if (typeof fmt === 'string' && !ASCII_NUMBER_FORMATS.has(fmt)) return 'all';
      }
      // Texts this record paints beyond what the rendered-text traversal
      // spells: a DATE / TIME picture's literals (its items are ASCII:
      // digits, English month/weekday names, AM/PM).
      const own: string[] = [];
      if (typeof record.fieldType === 'string' && typeof record.instruction === 'string') {
        const format = parseFieldFormatSwitch(record.instruction);
        if (format !== null && !ASCII_NUMBER_FORMATS.has(format)) return 'all';
        if (record.fieldType === 'date' || record.fieldType === 'time') {
          const picture = parseDateTimePictureSwitch(record.instruction);
          if (picture) own.push(picture);
        }
      }
      // §17.3.3.25 ruby: the guide text paints above its base run.
      const ruby = record.ruby as { text?: unknown } | null | undefined;
      if (ruby && typeof ruby === 'object' && typeof ruby.text === 'string') own.push(ruby.text);
      texts.push(...own);
      // §17.3.2.5 caps / §17.3.2.33 smallCaps: the shaper paints the
      // upper-cased display text of the run (segment-builder.ts), including
      // a field's synthesized result, so every text the run paints is also
      // demanded upper-cased.
      if (record.allCaps === true || record.smallCaps === true) {
        for (const value of [record.text, record.fallbackText, ...own]) {
          if (typeof value === 'string') texts.push(value.toUpperCase());
        }
      }
      if (typeof record.emphasisMark === 'string') texts.push(emphasisGlyph(record.emphasisMark));
    }
    for (const child of Object.values(value)) {
      if (child !== null && typeof child === 'object') pending.push(child);
    }
  }
  return texts;
}

/**
 * The Unicode scalars a document paints, for `unicode-range` subset laziness
 * (core `scalar-demand.ts` owns the scalar rules: PUA symbol codes outside
 * a known face list and unknown scalars decline to `'all'`). Texts come from
 * {@link docxRenderedTextUsages} (body, tables, headers/footers, notes,
 * text-box stories, numbering markers, field results), plus
 * {@link LAYOUT_SYNTHESIZED_SCALARS} and the synthesised paint of
 * {@link docxSynthesizedPaint}. A usage in a symbol font (Symbol /
 * Wingdings, including `w:sym` runs and symbol bullets) may be remapped to
 * other Unicode scalars by the painter (core `symbol-font.ts`), so it demands
 * every scalar that remapping may paint (`addSymbolText`). The set is bounded by `HARD_MAX_FONT_DEMAND_SCALARS`; beyond
 * it the answer is `'all'`. `'all'` is the conservative answer, never an error.
 */
export function docxFontDemandScalars(doc: DocxDocumentModel): FontDemandScalars {
  const synthesized = docxSynthesizedPaint(doc);
  if (synthesized === 'all') return 'all';
  // Computed in the realm that loads the faces: nothing is serialized, so the
  // only bound is the scalar-count ceiling inside the collector.
  const collector = new FontDemandScalarCollector({ budget: Number.POSITIVE_INFINITY });
  if (!collector.add(String.fromCodePoint(...LAYOUT_SYNTHESIZED_SCALARS))) return 'all';
  // Every symbol-encoded family the document paints with. A synthesized text
  // (a caps run's upper case, a field picture's literals, a field result) is painted in its
  // run's fonts, which the synthesized list does not carry, so it is
  // collected as if painted in each of them: a finite over-approximation.
  const symbolFamilies = new Set<string>();
  for (const usage of docxRenderedTextUsages(doc, 'font-demand')) {
    for (const family of usage.fontFamilies) if (isSymbolFontFamily(family)) symbolFamilies.add(family!);
    // The usage lists every slot face the run's symbol conversion consults
    // (segment-builder symbolSegments, numbering-marker): with none
    // symbol-encoded, a PUA code paints as itself.
    if (!collector.addSymbolText(usage.text, usage.fontFamilies, true)) return 'all';
  }
  // Layout-generated text (a field result such as a PAGE number) is painted
  // in its field's fonts too.
  if (!collector.addSymbolText(String.fromCodePoint(...LAYOUT_SYNTHESIZED_SCALARS), symbolFamilies)) return 'all';
  for (const text of synthesized) if (!collector.addSymbolText(text, symbolFamilies)) return 'all';
  return collector.result();
}

/** Whether the document names a WMF/EMF picture part (by type or part
 * name), anywhere in the parsed model. */
function docxReferencesMetafile(doc: DocxDocumentModel): boolean {
  const pending: unknown[] = [doc];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value) || ArrayBuffer.isView(value)) continue;
    seen.add(value);
    const record = value as Record<string, unknown>;
    if ((typeof record.mimeType === 'string' && isMetafileMime(record.mimeType))
      || (typeof record.imagePath === 'string' && METAFILE_PATH.test(record.imagePath))) return true;
    for (const child of Object.values(record)) if (child !== null && typeof child === 'object') pending.push(child);
  }
  return false;
}

/** Metafile pictures demand their text's scalars when they play (core
 * `bindMetafileFontDemand`). */
export { metafileFontDemand } from '@silurus/ooxml-core/internal/metafile-font-demand';

/** Loaded font sources for layout, plus the session that owns their faces. */
export interface DocxLoadedFontSources extends DocxSourceFonts {
  readonly session: FontSourceSession;
}

/**
 * Resolve, register and load one document's faces into `fontSet` (the
 * painting realm's set) before pagination measures text. Rejects with an
 * `AbortError` when `signal` aborts; a failed family or face (including a
 * run-time limit breach) only warns and keeps its host fallback. The caller owns `session` and releases it with the
 * document.
 */
export async function loadDocxFontSources(
  doc: DocxDocumentModel,
  plan: DocxFontFamilyPlan,
  sources: readonly FontSourceDescriptor[],
  // Worker mode passes the bridge's session options: `fontResolver`
  // functions run on main, and main fetches their claimed URL faces (core
  // resolver-bridge.ts).
  options: Omit<FontSourceSessionOptions, 'sources' | 'resolverFailures' | 'readMetricTables'> = {},
): Promise<DocxLoadedFontSources> {
  // DOCX lays out with the served faces' own tables (source-font-metrics.ts):
  // the session reads them before the snapshot. PPTX and XLSX do not.
  const session = new FontSourceSession({
    ...options, sources, readMetricTables: readFontResourceMetrics,
    // Without an `'all'` paint path (docxSynthesizedPaint: no chart or other
    // unmodelled text) every painted run is regular or bold:
    // WordprocessingML §17.3.2.1 `b` (layout paragraph.ts, tabs,
    // shape-drawing-plan text paths), so 400 and 700 are the only weights.
    // The session uses them only for a finite demand (an `'all'` demand
    // registers every declaration as its own face). Metafile text paints at
    // its LOGFONT weight (MS-WMF 2.2.1.2, any of 1–1000), which only the
    // picture's records tell, so a document with metafile pictures registers
    // every declaration as its own face too.
    ...(docxReferencesMetafile(doc) ? {} : { requestedWeights: DOCX_RUN_WEIGHTS }),
  });
  try {
    // The scalars are collected while the families' stylesheets download.
    const loaded = await session.ensureWithScalars({
      families: plan.families,
      fallbackFamilies: plan.fallbackFamilies,
      substituteScripts: plan.substituteScripts,
    }, () => docxFontDemandScalars(doc));
    return {
      session,
      names: Object.freeze([...plan.families, ...plan.fallbackFamilies]),
      loaded,
      lineMetrics: servedFontLineMetrics,
    };
  } catch (error) {
    session.release();
    throw error;
  }
}
