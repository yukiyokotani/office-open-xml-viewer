import type { DocxDocumentModel } from './types.js';

const DOM_VERTICAL_PAGE_DIRECTIONS = new Set(['tbRl', 'tbRlV', 'tbLrV']);

/**
 * Whether correct retained glyph acquisition may require the DOM-only Canvas
 * OpenType `vert` route. OffscreenCanvas exposes neither an equivalent feature
 * switch nor a way to prove the selected font's vertical alternate, so such a
 * document must not silently take the worker renderer's geometric fallback.
 *
 * Vertical sections can occur in the body or in a later section transition.
 * Walking the immutable parser model keeps this capability decision independent
 * of where the section property was authored. Text-box textVert is intentionally
 * excluded: that layout path does not yet use the DOM vertical-glyph planner.
 */
export function documentRequiresDomVerticalGlyphLayout(
  document: DocxDocumentModel,
): boolean {
  const pending: unknown[] = [document];
  const seen = new Set<object>();
  while (pending.length > 0) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value)) continue;
    seen.add(value);
    if (!Array.isArray(value)) {
      const record = value as Record<string, unknown>;
      if (
        typeof record.textDirection === 'string'
        && DOM_VERTICAL_PAGE_DIRECTIONS.has(record.textDirection)
      ) return true;
    }
    // A loop, not a spread: a value can hold huge arrays (VML text
    // rectangles), whose spread would overflow the call stack.
    for (const child of Object.values(value)) pending.push(child);
  }
  return false;
}
