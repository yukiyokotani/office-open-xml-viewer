import type { DocxDocumentModel } from './types.js';

// Run languages. Ruby guides are shaped without a Canvas language
// (rubyGuideShapeRequest), so their own tags are not consulted.
const LANGUAGE_KEYS = new Set(['lang', 'langEastAsia', 'langBidi']);

/**
 * Whether the document names any §17.3.2.20 run language (other than Word's
 * `x-none`). Such text reaches Canvas `lang` wherever the measuring realm can
 * prove named-family ownership (core createNamedFamilyOracle), so a render
 * worker lays it out only where that capability holds; a document without one
 * never consults it and keeps the worker whatever the engine supports.
 */
export function documentCarriesRunLanguage(document: DocxDocumentModel): boolean {
  const pending: unknown[] = [document];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value)) continue;
    seen.add(value);
    if (Array.isArray(value)) {
      // One push per child: spreading a wide array as call arguments would
      // exceed the engine's argument limit on large documents.
      for (const child of value) pending.push(child);
      continue;
    }
    for (const [key, field] of Object.entries(value as Record<string, unknown>)) {
      if (LANGUAGE_KEYS.has(key) && isLanguageValue(field)) return true;
      pending.push(field);
    }
  }
  return false;
}

/** A language string, or a typography wire value carrying one. */
function isLanguageValue(field: unknown): boolean {
  const value = field !== null && typeof field === 'object' && 'value' in field
    ? (field as { value: unknown }).value : field;
  return typeof value === 'string' && value !== '' && value.toLowerCase() !== 'x-none';
}
