import { canvasLanguageTag } from './layout/text.js';

// Run languages. Ruby guides are shaped without a Canvas language
// (rubyGuideShapeRequest), so their own tags are not consulted.
const LANGUAGE_KEYS = new Set(['lang', 'langEastAsia', 'langBidi', 'langScriptClass']);

/**
 * Whether a model field names a §17.3.2.20 run language (other than Word's
 * `x-none`). Such text reaches Canvas `lang` wherever the measuring realm can
 * prove named-family ownership (core createNamedFamilyOracle), so a render
 * worker lays a document carrying one out only where that capability holds;
 * a document without one never consults it and keeps the worker whatever the
 * engine supports. The document-wide fact is collected by the one capability
 * walk in vertical-render-capability.ts (documentRenderCapabilities).
 */
export function isRunLanguageField(key: string, field: unknown): boolean {
  if (!LANGUAGE_KEYS.has(key)) return false;
  const value = field !== null && typeof field === 'object' && 'value' in field
    ? (field as { value: unknown }).value : field;
  // The tag shaping would admit (x-none, ill-formed and extlang tags name none).
  return typeof value === 'string' && canvasLanguageTag(value) !== undefined;
}
