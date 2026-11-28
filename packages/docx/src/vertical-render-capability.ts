import { isRunLanguageField } from './run-language-capability.js';
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

/**
 * Both capability facts from one walk of the immutable parser model, so a
 * load pays a single traversal whatever it needs to know; the walk stops once
 * both facts hold. Walking the model keeps the decisions independent of where
 * a section property or a run language was authored.
 *
 * The parser model is a tree, so the walk retains no visited set: only the
 * objects on the current path are held (cycle protection for caller-owned
 * compatibility models), and a subtree shared by several parents is simply
 * visited once per parent.
 */
export function documentRenderCapabilities(document: DocxDocumentModel): DocumentRenderCapabilities {
  let requiresDomVerticalGlyphLayout = false;
  let carriesRunLanguage = false;
  // Parallel stacks: a value to visit, or (exit) an object whose subtree is done.
  const pending: object[] = [document];
  const exits: boolean[] = [false];
  const path = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop()!;
    if (exits.pop()) {
      path.delete(value);
      continue;
    }
    if (path.has(value)) continue;
    path.add(value);
    pending.push(value);
    exits.push(true);
    if (Array.isArray(value)) {
      // One push per child: spreading a wide array as call arguments would
      // exceed the engine's argument limit on large documents.
      for (const child of value) {
        if (child !== null && typeof child === 'object') {
          pending.push(child);
          exits.push(false);
        }
      }
      continue;
    }
    const record = value as Record<string, unknown>;
    for (const key of Object.keys(record)) {
      const field = record[key];
      if (!requiresDomVerticalGlyphLayout && key === 'textDirection'
        && typeof field === 'string' && DOM_VERTICAL_PAGE_DIRECTIONS.has(field)) {
        requiresDomVerticalGlyphLayout = true;
      }
      if (!carriesRunLanguage && isRunLanguageField(key, field)) carriesRunLanguage = true;
      if (requiresDomVerticalGlyphLayout && carriesRunLanguage) {
        return { requiresDomVerticalGlyphLayout, carriesRunLanguage };
      }
      if (field !== null && typeof field === 'object') {
        pending.push(field);
        exits.push(false);
      }
    }
  }
  return { requiresDomVerticalGlyphLayout, carriesRunLanguage };
}
