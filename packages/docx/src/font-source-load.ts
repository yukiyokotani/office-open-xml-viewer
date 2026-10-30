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
  type FontDemandScalars,
} from '@silurus/ooxml-core/internal/font-scalar-demand';
import {
  FontSourceSession,
  type FontSessionDemand,
  type FontResolverRunner,
  type FontSource,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import {
  isMetafileMime,
  isSymbolFontFamily,
  parseDateTimePictureSwitch,
  parseFieldFormatSwitch,
} from '@silurus/ooxml-core';
import { docxRenderedTextUsages } from './document-content.js';
import type { DocxFontFamilyPlan } from './font-plan.js';
import { emphasisGlyph } from './layout/paragraph.js';
import { LAYOUT_SYNTHESIZED_SCALARS } from './line-layout.js';
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

/** §17.18.93 ST_TextDirection: only `lrTb` is horizontal. */
const HORIZONTAL_TEXT_DIRECTIONS: ReadonlySet<string> = new Set(['lrTb']);

const METAFILE_PATH = /\.(?:wmf|emf|wmz|emz)$/i;

/**
 * Paint facts the rendered-text traversal does not spell, from one walk over
 * the parsed model (the same generic walk `vertical-render-capability.ts`
 * uses). Returns the extra texts layout synthesises for this document, or
 * `'all'` for a paint path whose scalars cannot be proven from the text:
 *
 * - vertical text (section, cell or text-box direction other than `lrTb` /
 *   `horz`; §17.3.2.10 `eastAsianLayout` vert/combine): vertical forms and
 *   rotated punctuation are font-selected alternates;
 * - ruby (§17.3.3.25), RTL runs and bidi paragraphs (UAX #9 mirroring paints
 *   scalars the text does not spell);
 * - charts and equations (their own text and number formatting) and WMF/EMF
 *   pictures (metafile text records paint with the document's fonts, and the
 *   Office-face web alias); a raster-labelled part whose bytes are a metafile
 *   is not detected (residual);
 * - page / note numbers in a non-ASCII number format.
 *
 * Synthesised texts: `allCaps` / `smallCaps` runs paint their text upper-cased
 * (§17.3.2.5/.33), emphasis marks paint their glyph (§17.3.2.12), and a DATE /
 * TIME picture (§17.16.4.1) paints its literal characters.
 */
function docxSynthesizedPaint(doc: DocxDocumentModel): 'all' | string[] {
  const texts: string[] = [];
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
      if (typeof record.textDirection === 'string' && !HORIZONTAL_TEXT_DIRECTIONS.has(record.textDirection)) return 'all';
      if (typeof record.textVert === 'string' && record.textVert !== 'horz') return 'all';
      if (record.ruby != null || record.eastAsianVert === true || record.eastAsianCombine === true) return 'all';
      if (record.rtl === true || record.bidi === true) return 'all';
      if (record.type === 'chart' || record.type === 'math') return 'all';
      if (typeof record.mimeType === 'string' && isMetafileMime(record.mimeType)) return 'all';
      if (typeof record.imagePath === 'string' && METAFILE_PATH.test(record.imagePath)) return 'all';
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
 * (core `scalar-demand.ts` owns the scalar rules: strong RTL, PUA symbol
 * codes and non-inert scalars decline to `'all'`). Texts come from
 * {@link docxRenderedTextUsages} (body, tables, headers/footers, notes,
 * text-box stories, numbering markers, field results), plus
 * {@link LAYOUT_SYNTHESIZED_SCALARS} and the synthesised paint of
 * {@link docxSynthesizedPaint}. Any usage whose family is a symbol font
 * (Symbol / Wingdings, including `w:sym` runs and symbol bullets) is remapped
 * to other Unicode scalars by the painter (core `symbol-font.ts`), so it makes
 * demand `'all'`. The set is bounded by `HARD_MAX_FONT_DEMAND_SCALARS`; beyond
 * it the answer is `'all'`. `'all'` is the conservative answer, never an error.
 */
export function docxFontDemandScalars(doc: DocxDocumentModel): FontDemandScalars {
  const synthesized = docxSynthesizedPaint(doc);
  if (synthesized === 'all') return 'all';
  // Computed in the realm that loads the faces: nothing is serialized, so the
  // only bound is the scalar-count ceiling inside the collector.
  const collector = new FontDemandScalarCollector({ budget: Number.POSITIVE_INFINITY });
  if (!collector.add(String.fromCodePoint(...LAYOUT_SYNTHESIZED_SCALARS))) return 'all';
  for (const text of synthesized) if (!collector.add(text)) return 'all';
  for (const usage of docxRenderedTextUsages(doc, 'font-demand')) {
    if (usage.fontFamilies.some((family) => isSymbolFontFamily(family?.trim()))) return 'all';
    if (!collector.add(usage.text)) return 'all';
  }
  return collector.result();
}

/** One document's demand: the family plan plus its scalars. */
export function docxFontDemand(doc: DocxDocumentModel, plan: DocxFontFamilyPlan): FontSessionDemand {
  return Object.freeze({
    families: plan.families,
    fallbackFamilies: plan.fallbackFamilies,
    scalars: docxFontDemandScalars(doc),
    substituteScripts: plan.substituteScripts,
  });
}

/** Loaded font sources for layout, plus the session that owns their faces. */
export interface DocxLoadedFontSources {
  readonly session: FontSourceSession;
  /** Demanded names (authored, then fallback) the inventory iterates. */
  readonly names: readonly string[];
  readonly loaded: LoadedFontSources;
}

/**
 * Resolve, register and load one document's faces into `fontSet` (the
 * painting realm's set) before pagination measures text. Rejects with
 * `OoxmlError('font-source-limit')` for a misconfigured or hostile source and
 * with an `AbortError` when `signal` aborts; a failed face only warns and
 * keeps its host fallback. The caller owns `session` and releases it with the
 * document.
 */
export async function loadDocxFontSources(
  doc: DocxDocumentModel,
  plan: DocxFontFamilyPlan,
  sources: readonly FontSource[],
  options: Readonly<{
    fontSet?: FontFaceSet | null;
    signal?: AbortSignal;
    runResolver?: FontResolverRunner;
    resolverCeilingOwnedByRunner?: boolean;
  }> = {},
): Promise<DocxLoadedFontSources> {
  const session = new FontSourceSession({
    sources,
    ...(options.fontSet === undefined ? {} : { fontSet: options.fontSet }),
    ...(options.signal ? { signal: options.signal } : {}),
    // Worker mode: `fontResolver` functions run on main (core resolver-bridge.ts).
    ...(options.runResolver ? { runResolver: options.runResolver } : {}),
    ...(options.resolverCeilingOwnedByRunner ? { resolverCeilingOwnedByRunner: true } : {}),
  });
  try {
    const loaded = await session.ensure(docxFontDemand(doc, plan));
    return { session, names: Object.freeze([...plan.families, ...plan.fallbackFamilies]), loaded };
  } catch (error) {
    session.release();
    throw error;
  }
}
