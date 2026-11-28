import { canvasLanguageAxes, scriptSlotLanguage, type FontScriptSlot } from './layout/text.js';

/** Run kinds whose objects carry §17.3.2.20 language facts directly. */
const LANGUAGE_RUN_TYPES = new Set(['text', 'field']);

/** Whether `record` is a run object (text or field), whose language
 * properties {@link recordNamesRunLanguage} may read. Any other object that
 * happens to own a property called `lang` (a font-table entry key, say) is
 * not a run language. */
export function isLanguageRunRecord(record: Record<string, unknown>): boolean {
  const type = record.type;
  return typeof type === 'string' && LANGUAGE_RUN_TYPES.has(type);
}

const unwrap = (field: unknown): string | undefined => {
  const value = field !== null && typeof field === 'object' && 'value' in field
    ? (field as { value: unknown }).value : field;
  return typeof value === 'string' ? value : undefined;
};

const LANGUAGE_SLOTS: readonly FontScriptSlot[] = ['ascii', 'eastAsia', 'complexScript'];

/**
 * Whether a run or run-property fact object names a §17.3.2.20 run language
 * that shaping would hand to Canvas `lang`. The decision is shaping's own
 * (layout/text.ts): an ODF script-class language takes precedence (empty
 * names none), a DOCX axis marked automatic names none, and x-none,
 * ill-formed and extlang tags name none. Such text reaches Canvas `lang`
 * wherever the measuring realm can prove named-family ownership (core
 * createNamedFamilyOracle), so a render worker lays a document carrying one
 * out only where that capability holds; a document without one never consults
 * it and keeps the worker whatever the engine supports. The document-wide
 * fact is collected by the one capability walk in vertical-render-capability.ts
 * (documentRenderCapabilities), which calls this only for run objects and
 * run-property facts.
 */
export function recordNamesRunLanguage(record: Record<string, unknown>): boolean {
  const language = unwrap(record.lang);
  const eastAsia = unwrap(record.langEastAsia);
  const bidi = unwrap(record.langBidi);
  const scriptClass = unwrap(record.langScriptClass);
  if (language === undefined && eastAsia === undefined && bidi === undefined && scriptClass === undefined) return false;
  const request = {
    language,
    eastAsiaLanguage: eastAsia,
    ...(scriptClass !== undefined ? { scriptClassLanguage: scriptClass } : {}),
    ...canvasLanguageAxes({
      langEastAsiaAutomatic: record.langEastAsiaAutomatic === true,
      langBidi: bidi,
      langBidiAutomatic: record.langBidiAutomatic === true,
    }),
  };
  return LANGUAGE_SLOTS.some((slot) => scriptSlotLanguage(request, slot) !== undefined);
}
