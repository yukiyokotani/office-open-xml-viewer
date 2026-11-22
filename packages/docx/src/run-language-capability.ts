import type { DocxDocumentModel } from './types.js';

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
    if (!Array.isArray(value)) {
      for (const [key, field] of Object.entries(value as Record<string, unknown>)) {
        if (LANGUAGE_KEYS.has(key) && typeof field === 'string' && field !== '' && field.toLowerCase() !== 'x-none') {
          return true;
        }
      }
    }
    pending.push(...Object.values(value));
  }
  return false;
}
