import { isLanguageRunRecord, LANGUAGE_FACT_KEYS, recordNamesRunLanguage } from './run-language-capability.js';
import type { DocxDocumentModel } from './types.js';

const DOM_VERTICAL_PAGE_DIRECTIONS = new Set(['tbRl', 'tbRlV', 'tbLrV']);

/** Document-wide facts that decide where a document can be laid out. */
export interface DocumentRenderCapabilities {
  /**
   * Whether correct retained glyph acquisition may require the DOM-only
   * Canvas OpenType `vert` route. OffscreenCanvas exposes neither an
   * equivalent feature switch nor a way to prove the selected font's vertical
   * alternate, so such a document must not silently take the worker
   * renderer's geometric fallback. Vertical sections can occur in the body or
   * in a later section transition. Text-box textVert is intentionally
   * excluded: that layout path does not yet use the DOM vertical-glyph
   * planner.
   */
  readonly requiresDomVerticalGlyphLayout: boolean;
  /** Whether any run names a language Canvas `lang` may take
   * (run-language-capability.ts). */
  readonly carriesRunLanguage: boolean;
}

/** The early answer of a walk asked to stop at vertical layout: the
 * document leaves the worker, so its language fact is not collected. */
export interface DomVerticalGlyphLayoutRequired {
  readonly requiresDomVerticalGlyphLayout: true;
  readonly carriesRunLanguage?: undefined;
}

/**
 * Both capability facts from one walk of the immutable parser model, so a
 * load pays a single traversal whatever it needs to know; the walk stops once
 * both facts hold. Walking the model keeps the decisions independent of where
 * a section property or a run language was authored.
 *
 * With `stopOnVerticalLayout` (render-worker admission) the walk returns as
 * soon as it finds a DOM-only vertical section, as the vertical-only walk did
 * before languages: such a document is laid out on the main thread, where the
 * layout source adapter walks the model for both facts itself.
 *
 * Each object is visited once (identity set): a subtree shared by several
 * parents in a caller-owned compatibility model costs one visit, and a cyclic
 * one terminates.
 */
export function documentRenderCapabilities(document: DocxDocumentModel): DocumentRenderCapabilities;
export function documentRenderCapabilities(
  document: DocxDocumentModel,
  options: { readonly stopOnVerticalLayout: true },
): DocumentRenderCapabilities | DomVerticalGlyphLayoutRequired;
export function documentRenderCapabilities(
  document: DocxDocumentModel,
  options?: { readonly stopOnVerticalLayout: true },
): DocumentRenderCapabilities | DomVerticalGlyphLayoutRequired {
  const stopOnVerticalLayout = options?.stopOnVerticalLayout === true;
  let requiresDomVerticalGlyphLayout = false;
  let carriesRunLanguage = false;
  const pending: object[] = [document];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop()!;
    if (seen.has(value)) continue;
    seen.add(value);
    if (Array.isArray(value)) {
      // One push per child: spreading a wide array as call arguments would
      // exceed the engine's argument limit on large documents.
      for (const child of value) {
        if (child !== null && typeof child === 'object' && !seen.has(child)) pending.push(child);
      }
      continue;
    }
    const record = value as Record<string, unknown>;
    // Only run objects carry a run language; a property merely named `lang`
    // elsewhere (a font-table key) must not route the document.
    if (!carriesRunLanguage && isLanguageRunRecord(record) && recordNamesRunLanguage(record)) {
      carriesRunLanguage = true;
    }
    for (const key of Object.keys(record)) {
      const field = record[key];
      if (!requiresDomVerticalGlyphLayout && key === 'textDirection'
        && typeof field === 'string' && DOM_VERTICAL_PAGE_DIRECTIONS.has(field)) {
        if (stopOnVerticalLayout) return { requiresDomVerticalGlyphLayout: true };
        requiresDomVerticalGlyphLayout = true;
      }
      if (!carriesRunLanguage && LANGUAGE_FACT_KEYS.has(key) && field !== null && typeof field === 'object'
        && !Array.isArray(field) && recordNamesRunLanguage(field as Record<string, unknown>)) {
        carriesRunLanguage = true;
      }
      if (requiresDomVerticalGlyphLayout && carriesRunLanguage) {
        return { requiresDomVerticalGlyphLayout, carriesRunLanguage };
      }
      if (field !== null && typeof field === 'object' && !seen.has(field)) pending.push(field);
    }
  }
  return { requiresDomVerticalGlyphLayout, carriesRunLanguage };
}
