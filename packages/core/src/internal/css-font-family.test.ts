import { describe, expect, it } from 'vitest';
import { scopeServedFontFamilies } from './css-font-family.js';

describe('scopeServedFontFamilies', () => {
  const families = new Map([['lato', 'ooxml-font-source-1-0'], ['noto sans jp', 'ooxml-font-source-1-1']]);

  it('names each served family\'s private family right before it, anywhere in the list', () => {
    expect(scopeServedFontFamilies('"Calibri", "Lato", Noto Sans JP, sans-serif', families))
      .toBe('"Calibri", "ooxml-font-source-1-0", "Lato", "ooxml-font-source-1-1", Noto Sans JP, sans-serif');
  });

  it('is idempotent and leaves lists without served families, generics and malformed lists alone', () => {
    const once = scopeServedFontFamilies('"LATO", serif', families);
    expect(once).toBe('"ooxml-font-source-1-0", "LATO", serif');
    expect(scopeServedFontFamilies(once, families)).toBe(once);
    expect(scopeServedFontFamilies('"Arial", sans-serif', families)).toBe('"Arial", sans-serif');
    expect(scopeServedFontFamilies('"Lato', families)).toBe('"Lato');
    expect(scopeServedFontFamilies('"Lato"', new Map())).toBe('"Lato"');
  });
});
