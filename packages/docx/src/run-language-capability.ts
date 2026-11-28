import { canvasLanguageTag } from './layout/text.js';

// Run languages. Ruby guides are shaped without a Canvas language
// (rubyGuideShapeRequest), so their own tags are not consulted.
const LANGUAGE_KEYS = new Set(['lang', 'langEastAsia', 'langBidi', 'langScriptClass']);

/** Run kinds whose objects carry §17.3.2.20 language facts directly. */
const LANGUAGE_RUN_TYPES = new Set(['text', 'field']);
/** Properties holding the run-property facts of a non-content glyph:
 * paragraph mark, numbering level, inline object host. */
export const LANGUAGE_FACT_KEYS = new Set(['paragraphMarkFontFacts', 'fontFacts', '__inlineObjectFontFacts']);

/** Whether `record` is a run object (text or field), whose language
 * properties {@link recordNamesRunLanguage} may read. Any other object that
 * happens to own a property called `lang` (a font-table entry key, say) is
 * not a run language. */
export function isLanguageRunRecord(record: Record<string, unknown>): boolean {
  const type = record.type;
  return typeof type === 'string' && LANGUAGE_RUN_TYPES.has(type);
}

/**
 * Whether a run or run-property fact object names a §17.3.2.20 run language
 * (other than Word's `x-none`). Such text reaches Canvas `lang` wherever the
 * measuring realm can prove named-family ownership (core
 * createNamedFamilyOracle), so a render worker lays a document carrying one
 * out only where that capability holds; a document without one never consults
 * it and keeps the worker whatever the engine supports. The document-wide
 * fact is collected by the one capability walk in vertical-render-capability.ts
 * (documentRenderCapabilities), which calls this only for run objects and
 * run-property facts.
 */
export function recordNamesRunLanguage(record: Record<string, unknown>): boolean {
  for (const key of LANGUAGE_KEYS) {
    const field = record[key];
    if (field === undefined) continue;
    const value = field !== null && typeof field === 'object' && 'value' in field
      ? (field as { value: unknown }).value : field;
    // The tag shaping would admit (x-none, ill-formed and extlang tags name none).
    if (typeof value === 'string' && canvasLanguageTag(value) !== undefined) return true;
  }
  return false;
}
