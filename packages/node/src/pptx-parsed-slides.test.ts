import { describe, expect, it } from 'vitest';
import { markOdfPptxTables } from './pptx-parsed-slides.ts';

describe('markOdfPptxTables', () => {
  it('marks slide tables and the fallback tables the renderer paints, however deep the nesting', () => {
    // 12,000 nested equation fallbacks: only the first level is ever painted
    // (renderSlide replaces an element by its own fallback elements), so the
    // walk stays at that level and cannot overflow the stack.
    type Element = { type: string; odfRowHeights?: boolean; mathFallback?: { elements: Element[]; before: number; after: number } };
    const deepest: Element = { type: 'table' };
    let nested: Element = deepest;
    for (let depth = 0; depth < 12_000; depth++) {
      nested = { type: 'table', mathFallback: { elements: [nested], before: 0, after: 0 } };
    }
    const top: Element = { type: 'shape', mathFallback: { elements: [nested], before: 0, after: 0 } };
    const plain: Element = { type: 'table' };
    const slide = { elements: [top, plain] };
    expect(() => markOdfPptxTables(slide)).not.toThrow();
    expect(plain.odfRowHeights).toBe(true);
    expect(nested.odfRowHeights).toBe(true);
    expect(top.odfRowHeights).toBeUndefined();
    expect(deepest.odfRowHeights).toBeUndefined();
  });
});
